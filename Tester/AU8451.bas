Attribute VB_Name = "AU8451"
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

Public Type msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Public Const WM_USER = &H400
Public Const WM_FT_TEST_READY = WM_USER + &H800

'SD card Test
Public Const WM_FT_SD_START = WM_USER + &H100
Public Const WM_FT_SD_PASS = WM_USER + &H120
Public Const WM_FT_SD_FAIL = WM_USER + &H140
Public Const WM_FT_SD_SPEED_DOWN = WM_USER + &H160
                                             
'CF Card Test
Public Const WM_FT_CF_START = WM_USER + &H200
Public Const WM_FT_CF_PASS = WM_USER + &H220
Public Const WM_FT_CF_FAIL = WM_USER + &H240
                                             
'XD Card Test
Public Const WM_FT_XD_START = WM_USER + &H300
Public Const WM_FT_XD_PASS = WM_USER + &H320
Public Const WM_FT_XD_FAIL = WM_USER + &H340
                                             
'MS Card Test
Public Const WM_FT_MS_START = WM_USER + &H400
Public Const WM_FT_MS_PASS = WM_USER + &H420
Public Const WM_FT_MS_FAIL = WM_USER + &H440

'uSD Card Test
Public Const WM_FT_uSD_START = WM_USER + &H600
Public Const WM_FT_uSD_PASS = WM_USER + &H620
Public Const WM_FT_uSD_FAIL = WM_USER + &H640

'M2 Card Test
Public Const WM_FT_M2_START = WM_USER + &H700
Public Const WM_FT_M2_PASS = WM_USER + &H720
Public Const WM_FT_M2_FAIL = WM_USER + &H740

Public Const WM_FT_ClearBtn_Click = WM_USER + &H500

Public Sub LoadAP_Click_AU8451()
Dim TimePass
Dim rt2
    
    ' find window
    winHwnd = FindWindow(vbNullString, "AU8451FT_Tool")
     
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(Tester.hwnd, "open", App.Path & "\LibUSB\AU8451\AU8451FT.exe", "", "", SW_SHOW)
    End If
    
    SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags
    

End Sub


Public Function AU8451_SD_Test() As Byte
Dim rt2
Dim EntryTime As Long
Dim PassingTime As Long
Dim mMsg As msg
    
    AU8451_SD_Test = 0
    
    winHwnd = FindWindow(vbNullString, "AU8451FT_Tool")
    
    EntryTime = Timer
    
    If winHwnd <> 0 Then
    
        rt2 = PostMessage(winHwnd, WM_FT_SD_START, 0&, 0&)
        
        Do
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
    
            PassTime = Timer - EntryTime

        Loop Until AlcorMPMessage = WM_FT_SD_PASS _
                Or AlcorMPMessage = WM_FT_SD_FAIL _
                Or AlcorMPMessage = WM_FT_SD_SPEED_DOWN _
                Or PassTime > 12
    
    End If
    
    If AlcorMPMessage = WM_FT_SD_SPEED_DOWN Then
        AU8451_SD_Test = 2
    ElseIf AlcorMPMessage = WM_FT_SD_PASS Then
        AU8451_SD_Test = 1
    End If

End Function

Public Function AU8451_uSD_Test() As Byte
Dim rt2
Dim EntryTime As Long
Dim PassingTime As Long
Dim mMsg As msg
    
    AU8451_uSD_Test = 0
    
    winHwnd = FindWindow(vbNullString, "AU8451FT_Tool")
    
    EntryTime = Timer
    
    If winHwnd <> 0 Then
    
        rt2 = PostMessage(winHwnd, WM_FT_uSD_START, 0&, 0&)
        
        Do
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
    
            PassTime = Timer - EntryTime

        Loop Until AlcorMPMessage = WM_FT_uSD_PASS _
                Or AlcorMPMessage = WM_FT_uSD_FAIL _
                Or PassTime > 4
    
    End If
    
    If AlcorMPMessage = WM_FT_uSD_PASS Then
        AU8451_uSD_Test = 1
    End If

End Function

Public Function AU8451_CF_Test() As Byte
Dim rt2
Dim EntryTime As Long
Dim PassingTime As Long
Dim mMsg As msg
    
    AU8451_CF_Test = 0
    
    winHwnd = FindWindow(vbNullString, "AU8451FT_Tool")
    
    EntryTime = Timer
    
    If winHwnd <> 0 Then
    
        rt2 = PostMessage(winHwnd, WM_FT_CF_START, 0&, 0&)
        
        Do
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
    
            PassTime = Timer - EntryTime

        Loop Until AlcorMPMessage = WM_FT_CF_PASS _
                Or AlcorMPMessage = WM_FT_CF_FAIL _
                Or PassTime > 4
    
    End If
    
    If AlcorMPMessage = WM_FT_CF_PASS Then
        AU8451_CF_Test = 1
    End If

End Function

Public Function AU8451_XD_Test() As Byte
Dim rt2
Dim EntryTime As Long
Dim PassingTime As Long
Dim mMsg As msg
    
    AU8451_XD_Test = 0
    
    winHwnd = FindWindow(vbNullString, "AU8451FT_Tool")
    
    EntryTime = Timer
    
    If winHwnd <> 0 Then
    
        rt2 = PostMessage(winHwnd, WM_FT_XD_START, 0&, 0&)
        
        Do
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
    
            PassTime = Timer - EntryTime

        Loop Until AlcorMPMessage = WM_FT_XD_PASS _
                Or AlcorMPMessage = WM_FT_XD_FAIL _
                Or PassTime > 4
    
    End If
    
    If AlcorMPMessage = WM_FT_XD_PASS Then
        AU8451_XD_Test = 1
    End If

End Function

Public Function AU8451_MS_Test() As Byte
Dim rt2
Dim EntryTime As Long
Dim PassingTime As Long
Dim mMsg As msg
    
    AU8451_MS_Test = 0
    
    winHwnd = FindWindow(vbNullString, "AU8451FT_Tool")
    
    EntryTime = Timer
    
    If winHwnd <> 0 Then
    
        rt2 = PostMessage(winHwnd, WM_FT_MS_START, 0&, 0&)
        
        Do
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
    
            PassTime = Timer - EntryTime

        Loop Until AlcorMPMessage = WM_FT_MS_PASS _
                Or AlcorMPMessage = WM_FT_MS_FAIL _
                Or PassTime > 4
    
    End If
    
    If AlcorMPMessage = WM_FT_MS_PASS Then
        AU8451_MS_Test = 1
    End If

End Function

Public Function AU8451_M2_Test() As Byte
Dim rt2
Dim EntryTime As Long
Dim PassingTime As Long
Dim mMsg As msg
    
    AU8451_M2_Test = 0
    
    winHwnd = FindWindow(vbNullString, "AU8451FT_Tool")
    
    EntryTime = Timer
    
    If winHwnd <> 0 Then
    
        rt2 = PostMessage(winHwnd, WM_FT_M2_START, 0&, 0&)
        
        Do
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
    
            PassTime = Timer - EntryTime

        Loop Until AlcorMPMessage = WM_FT_M2_PASS _
                Or AlcorMPMessage = WM_FT_M2_FAIL _
                Or PassTime > 4
    
    End If
    
    If AlcorMPMessage = WM_FT_M2_PASS Then
        AU8451_M2_Test = 1
    End If

End Function

Public Sub Clear_Btn_Click_AU8451()
Dim rt2
Dim EntryTime As Long
Dim PassingTime As Long
Dim mMsg As msg
    
    winHwnd = FindWindow(vbNullString, "AU8451FT_Tool")
    
    If winHwnd <> 0 Then
        Do
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
    
            PassTime = Timer - EntryTime
    
        Loop Until AlcorMPMessage = WM_FT_TEST_READY _
                Or PassTime > 1
    
        rt2 = PostMessage(winHwnd, WM_FT_ClearBtn_Click, 0&, 0&)
    End If

End Sub


Public Sub CloseAU8451FT_Tool()
Dim rt2 As Long
Dim EntryTime As Long
Dim PassingTime As Long
Dim mMsg As msg

    winHwnd = FindWindow(vbNullString, "AU8451FT_Tool")
    
    EntryTime = Timer
    
    If winHwnd <> 0 Then
        Do
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
    
            PassTime = Timer - EntryTime

        Loop Until AlcorMPMessage = WM_FT_TEST_READY _
                Or PassTime > 1
        
        Do
            rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
            Call MsecDelay(0.5)
            winHwnd = FindWindow(vbNullString, "AU8451FT_Tool")
        Loop While winHwnd <> 0
    End If
    
End Sub

Public Sub AU8451FTTestSub()

    If ChipName = "AU8451DBF22" Then
        Call AU8451DBF22TestSub
    ElseIf ChipName = "AU8451BBF22" Then
        Call AU8451BBF22TestSub
    ElseIf ChipName = "AU8451EBF22" Then
        Call AU8451EBF22TestSub
    End If


End Sub

Public Sub AU8451BBF22TestSub()
 
Dim OldTimer
Dim PassTime
Dim rt2
Dim mMsg As msg
Dim TempRes As Integer
Dim TempStr As String
Dim LightOff As Long
Dim LightOn As Long
Dim LightCount As Integer


'AU8451NBB 64TQ CON SOCKET V1
'==============================
'P1A   8 7 6 5  4 3 2 1
'      | | | |  | | | |
'            M  X C S E
'            S  D F D N
'                     A
'
'==============================


rv0 = 0 'Enum
rv1 = 0 'XD
rv2 = 0 'SD
rv3 = 0 'MS
rv4 = 0 'CF



If PCI7248InitFinish = 0 Then
    PCI7248Exist
End If

CardResult = DO_WritePort(card, Channel_P1A, &HFE)
Call MsecDelay(0.2)

AlcorMPMessage = 0


'===================== Wait AP Ready =====================

winHwnd = FindWindow(vbNullString, "AU8451FT_Tool")
If winHwnd = 0 Then

    Call LoadAP_Click_AU8451
    OldTimer = Timer

    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If

        PassTime = Timer - OldTimer

    Loop Until AlcorMPMessage = WM_FT_TEST_READY _
    Or PassTime > 5
    Tester.Print "Ready Time="; PassTime

    If PassTime > 5 Then
        TestResult = "Bin3"
        Tester.Label2.Caption = "Bin3"
        Tester.Label2.BackColor = RGB(255, 0, 0)
        Tester.Label9.BackColor = RGB(255, 0, 0)
        Call CloseAU8451FT_Tool
        CardResult = DO_WritePort(card, Channel_P1A, &HFF)
        Call MsecDelay(0.2)
        Exit Sub
    End If

End If

Call Clear_Btn_Click_AU8451

'===================== Find Device =====================

rv0 = WaitDevOn("pid_8431")
Call NewLabelMenu(0, "UnknowDevice", rv0, 1)

If rv0 <> 1 Then
    GoTo AU8451ResultLabel
End If


'===================== Get LED(OFF) Value  =====================
LightCount = 0

Do
    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
    LightCount = LightCount + 1
    LightOff = (LightOff And 1)
    Call MsecDelay(0.1)
Loop While (LightOff <> &H1) And (LightCount < 10)


'===================== Connect XD & Test =====================
CardResult = DO_WritePort(card, Channel_P1A, &HF6)
Call MsecDelay(0.2)

rv1 = AU8451_XD_Test()
Call NewLabelMenu(1, "XD Card", rv1, rv0)
If rv1 <> 1 Then
    GoTo AU8451ResultLabel
End If


'===================== Get LED(ON) Value & LED Result  =====================
LightCount = 0

Do
    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
    LightCount = LightCount + 1
    LightOn = (LightOn And 1)
    Call MsecDelay(0.1)
Loop While (LightOn <> &H0) And (LightCount < 10)

If (LightOn <> &H0) Or (LightOff <> &H1) Then
    UsbSpeedTestResult = GPO_FAIL
    rv2 = 2
    Call NewLabelMenu(1, "LED", rv1, rv0)
    GoTo AU8451ResultLabel
End If


'===================== Connect SD & Test =====================
CardResult = DO_WritePort(card, Channel_P1A, &HFC)
Call MsecDelay(0.2)

rv2 = AU8451_SD_Test()
Call NewLabelMenu(2, "SD Card", rv2, rv1)
If rv2 <> 1 Then
    GoTo AU8451ResultLabel
End If


'===================== Connect MS & Test =====================
CardResult = DO_WritePort(card, Channel_P1A, &HEE)
Call MsecDelay(0.2)

rv3 = AU8451_MS_Test()
Call NewLabelMenu(3, "MS Card", rv3, rv2)
If rv3 <> 1 Then
    GoTo AU8451ResultLabel
End If


'===================== Connect CF & Test =====================
CardResult = DO_WritePort(card, Channel_P1A, &HFA)
Call MsecDelay(0.2)

rv4 = AU8451_CF_Test()
Call NewLabelMenu(4, "CF Card", rv4, rv3)
If rv4 <> 1 Then
    GoTo AU8451ResultLabel
End If



AU8451ResultLabel:
    
    'Call Clear_Btn_Click_AU8451
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
    WaitDevOFF ("pid_8431")
    
    
    If (rv0 <> 1) Then
        TestResult = "Bin2"     'Unknow
        
    ElseIf (rv1 <> 1) Then
        TestResult = "Bin4"     'XD Fail
        
    ElseIf (rv2 <> 1) Then
        TestResult = "Bin3"     'SD Fail / LED Fail
        
    ElseIf (rv3 <> 1) Then
        TestResult = "Bin5"     'MS Fail
        
    ElseIf (rv4 <> 1) Then
        TestResult = "Bin3"     'CF Fail
        
    ElseIf (rv0 * rv1 * rv2 * rv3 * rv4 = 1) Then
        TestResult = "PASS"
        
    End If
    
    If TestResult <> "PASS" Then
        Call CloseAU8451FT_Tool
    End If
    

Call MsecDelay(0.2)

End Sub

Public Sub AU8451DBF22TestSub()
 
Dim OldTimer
Dim PassTime
Dim rt2
Dim mMsg As msg
Dim TempRes As Integer
Dim TempStr As String
Dim LightOff As Long
Dim LightOn As Long
Dim LightCount As Integer


'AU8431-NDB 48TQ CON SOCKET V1.02B
'==============================
'P1A   8 7 6 5  4 3 2 1
'      | | | |  | | | |
'            M  X   S E
'            S  D   D N
'                     A
'
'==============================


rv0 = 0 'Enum
rv1 = 0 'XD
rv2 = 0 'SD
rv3 = 0 'MS


If PCI7248InitFinish = 0 Then
    PCI7248Exist
End If

CardResult = DO_WritePort(card, Channel_P1A, &HFE)
Call MsecDelay(0.2)

AlcorMPMessage = 0


'===================== Wait AP Ready =====================

winHwnd = FindWindow(vbNullString, "AU8451FT_Tool")
If winHwnd = 0 Then

    Call LoadAP_Click_AU8451
    OldTimer = Timer

    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If

        PassTime = Timer - OldTimer

    Loop Until AlcorMPMessage = WM_FT_TEST_READY _
    Or PassTime > 5
    Tester.Print "Ready Time="; PassTime

    If PassTime > 5 Then
        TestResult = "Bin3"
        Tester.Label2.Caption = "Bin3"
        Tester.Label2.BackColor = RGB(255, 0, 0)
        Tester.Label9.BackColor = RGB(255, 0, 0)
        Call CloseAU8451FT_Tool
        CardResult = DO_WritePort(card, Channel_P1A, &HFF)
        Call MsecDelay(0.2)
        Exit Sub
    End If

End If

Call Clear_Btn_Click_AU8451

'===================== Find Device =====================

rv0 = WaitDevOn("pid_8431")
Call NewLabelMenu(0, "UnknowDevice", rv0, 1)

If rv0 <> 1 Then
    GoTo AU8451ResultLabel
End If


'===================== Get LED(OFF) Value  =====================
LightCount = 0

Do
    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
    LightCount = LightCount + 1
    LightOff = (LightOff And 1)
    Call MsecDelay(0.1)
Loop While (LightOff <> &H1) And (LightCount < 10)


'===================== Connect XD & Test =====================
CardResult = DO_WritePort(card, Channel_P1A, &HF6)
Call MsecDelay(0.2)

rv1 = AU8451_XD_Test()
Call NewLabelMenu(1, "XD Card", rv1, rv0)
If rv1 <> 1 Then
    GoTo AU8451ResultLabel
End If


'===================== Get LED(ON) Value & LED Result  =====================
LightCount = 0

Do
    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
    LightCount = LightCount + 1
    LightOn = (LightOn And 1)
    Call MsecDelay(0.1)
Loop While (LightOn <> &H0) And (LightCount < 10)

If (LightOn <> &H0) Or (LightOff <> &H1) Then
    UsbSpeedTestResult = GPO_FAIL
    rv2 = 2
    Call NewLabelMenu(1, "LED", rv1, rv0)
    GoTo AU8451ResultLabel
End If


'===================== Connect SD & Test =====================
CardResult = DO_WritePort(card, Channel_P1A, &HFC)
Call MsecDelay(0.2)

rv2 = AU8451_SD_Test()
Call NewLabelMenu(2, "SD Card", rv2, rv1)
If rv2 <> 1 Then
    GoTo AU8451ResultLabel
End If


'===================== Connect MS & Test =====================
CardResult = DO_WritePort(card, Channel_P1A, &HEE)
Call MsecDelay(0.2)

rv3 = AU8451_MS_Test()
Call NewLabelMenu(3, "MS Card", rv3, rv2)
If rv3 <> 1 Then
    GoTo AU8451ResultLabel
End If



AU8451ResultLabel:
    
    'Call Clear_Btn_Click_AU8451
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
    WaitDevOFF ("pid_8431")
    
    
    If (rv0 <> 1) Then
        TestResult = "Bin2"     'Unknow
        
    ElseIf (rv1 <> 1) Then
        TestResult = "Bin4"     'XD Fail
        
    ElseIf (rv2 = 0) Then
        TestResult = "Bin3"     'SD Fail / LED Fail
        
    ElseIf (rv2 = 2) Then
        TestResult = "Bin4"     'SD Fail / LED Fail
        
    ElseIf (rv3 <> 1) Then
        TestResult = "Bin5"     'MS Fail
        
    ElseIf (rv0 * rv1 * rv2 * rv3 = 1) Then
        TestResult = "PASS"
        
    End If
    
    If TestResult <> "PASS" Then
        Call CloseAU8451FT_Tool
    End If
    

Call MsecDelay(0.2)

End Sub

Public Sub AU8451EBF22TestSub()
 
Dim OldTimer
Dim PassTime
Dim rt2
Dim mMsg As msg
Dim TempRes As Integer
Dim TempStr As String
Dim LightOff As Long
Dim LightOn As Long
Dim LightCount As Integer


'AU8431-NDB 48TQ CON SOCKET V1.02B
'==============================
'P1A   8 7 6 5  4 3 2 1
'      | | | |  | | | |
'            M  S   S E
'            S  D   D N
'                   H A
'                   C
'
'==============================


rv0 = 0 'Enum
rv1 = 0 'SDHC
rv2 = 0 'uSD
rv3 = 0 'MS


If PCI7248InitFinish = 0 Then
    PCI7248Exist
End If

CardResult = DO_WritePort(card, Channel_P1A, &HFE)
Call MsecDelay(0.2)

AlcorMPMessage = 0


'===================== Wait AP Ready =====================

winHwnd = FindWindow(vbNullString, "AU8451FT_Tool")
If winHwnd = 0 Then

    Call LoadAP_Click_AU8451
    OldTimer = Timer

    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If

        PassTime = Timer - OldTimer

    Loop Until AlcorMPMessage = WM_FT_TEST_READY _
    Or PassTime > 5
    Tester.Print "Ready Time="; PassTime

    If PassTime > 5 Then
        TestResult = "Bin3"
        Tester.Label2.Caption = "Bin3"
        Tester.Label2.BackColor = RGB(255, 0, 0)
        Tester.Label9.BackColor = RGB(255, 0, 0)
        Call CloseAU8451FT_Tool
        CardResult = DO_WritePort(card, Channel_P1A, &HFF)
        Call MsecDelay(0.2)
        Exit Sub
    End If

End If

Call Clear_Btn_Click_AU8451

'===================== Find Device =====================

rv0 = WaitDevOn("pid_8431")
Call NewLabelMenu(0, "UnknowDevice", rv0, 1)

If rv0 <> 1 Then
    GoTo AU8451ResultLabel
End If


'===================== Get LED(OFF) Value  =====================
LightCount = 0

Do
    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
    LightCount = LightCount + 1
    LightOff = (LightOff And 1)
    Call MsecDelay(0.1)
Loop While (LightOff <> &H1) And (LightCount < 10)


'===================== Connect SDHC / uSD & Test =====================

CardResult = DO_WritePort(card, Channel_P1A, &HF4)
Call MsecDelay(0.2)

'SDHC Item
rv1 = AU8451_SD_Test()
Call NewLabelMenu(1, "SDXC Card", rv1, rv0)
If rv1 <> 1 Then
    GoTo AU8451ResultLabel
End If


'===================== Get LED(ON) Value & LED Result  =====================
LightCount = 0

Do
    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
    LightCount = LightCount + 1
    LightOn = (LightOn And 1)
    Call MsecDelay(0.1)
Loop While (LightOn <> &H0) And (LightCount < 10)

If (LightOn <> &H0) Or (LightOff <> &H1) Then
    UsbSpeedTestResult = GPO_FAIL
    rv1 = 2
    Call NewLabelMenu(1, "LED", rv1, rv0)
    GoTo AU8451ResultLabel
End If


'uSD Item
rv2 = AU8451_uSD_Test()
Call NewLabelMenu(2, "SD Card", rv2, rv1)
If rv2 <> 1 Then
    GoTo AU8451ResultLabel
End If



'===================== Connect MS & Test =====================
CardResult = DO_WritePort(card, Channel_P1A, &HEE)
Call MsecDelay(0.2)

rv3 = AU8451_M2_Test()
Call NewLabelMenu(3, "MS Card", rv3, rv2)
If rv3 <> 1 Then
    GoTo AU8451ResultLabel
End If



AU8451ResultLabel:
    
    'Call Clear_Btn_Click_AU8451
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
    WaitDevOFF ("pid_8431")
    
    
    If (rv0 <> 1) Then
        TestResult = "Bin2"     'Unknow
        
    ElseIf (rv1 <> 1) Then
        TestResult = "Bin3"     'SDHC Fail / LED Fail
        
    ElseIf (rv2 <> 1) Then
        TestResult = "Bin4"     'SD Fail
        
    ElseIf (rv3 <> 1) Then
        TestResult = "Bin5"     'MS Fail
        
    ElseIf (rv0 * rv1 * rv2 * rv3 = 1) Then
        TestResult = "PASS"
        
    End If
    
    If TestResult <> "PASS" Then
        Call CloseAU8451FT_Tool
    End If
    

Call MsecDelay(0.2)

End Sub
