Attribute VB_Name = "AU6485FT"
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
Public Const WM_FT_SPEED_ERROR = WM_USER + &H200
Public Const WM_FT_SetNBMD = WM_USER + &H520
Public Const WM_FT_SetNorMD = WM_USER + &H540


'SD card Test
Public Const WM_FT_SD_START = WM_USER + &H100
Public Const WM_FT_SD_PASS = WM_USER + &H120
Public Const WM_FT_SD_FAIL = WM_USER + &H140
Public Const WM_FT_SD_SPEED_DOWN = WM_USER + &H160
                                             
'MS Card Test
Public Const WM_FT_MS_START = WM_USER + &H400
Public Const WM_FT_MS_PASS = WM_USER + &H420
Public Const WM_FT_MS_FAIL = WM_USER + &H440

Public Const WM_FT_ClearBtn_Click = WM_USER + &H500

Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long

Const TH32CS_SNAPPROCESS = 2
Const MAX_PATH = 260

Private Type PROCESSENTRY32
    dwSize               As Long
    cntUsage             As Long
    th32ProcessID        As Long
    th32DefaultHeapID    As Long
    th32ModuleID         As Long
    cntThreads           As Long
    th32ParentProcessID  As Long
    pcPriClassBase       As Long
    dwFlags              As Long
    szexeFile            As String * MAX_PATH
End Type

Public Sub KillProcess(NameProcess As String)

Const PROCESS_ALL_ACCESS = &H1F0FFF
Const TH32CS_SNAPPROCESS As Long = 2&
Dim uProcess  As PROCESSENTRY32
Dim RProcessFound As Long
Dim hSnapshot As Long
Dim SzExename As String
Dim ExitCode As Long
Dim MyProcess As Long
Dim AppKill As Boolean
Dim AppCount As Integer
Dim i As Integer
Dim WinDirEnv As String
        
    If NameProcess <> "" Then
        AppCount = 0

        uProcess.dwSize = Len(uProcess)
        hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
        RProcessFound = ProcessFirst(hSnapshot, uProcess)
  
        Do
            i = InStr(1, uProcess.szexeFile, Chr(0))
            SzExename = LCase$(Left$(uProcess.szexeFile, i - 1))
            WinDirEnv = Environ("Windir") + "\"
            WinDirEnv = LCase$(WinDirEnv)
        
            If Right$(SzExename, Len(NameProcess)) = LCase$(NameProcess) Then
                AppCount = AppCount + 1
                MyProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
                AppKill = TerminateProcess(MyProcess, ExitCode)
                Call CloseHandle(MyProcess)
            End If
            RProcessFound = ProcessNext(hSnapshot, uProcess)
        Loop While RProcessFound
        Call CloseHandle(hSnapshot)
        
    End If

End Sub

Public Sub LoadAP_Click_AU6485()
Dim TimePass
Dim rt2
    
    ' find window
    winHwnd = FindWindow(vbNullString, "AU6485FT_Tool")
     
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(Tester.hwnd, "open", App.Path & "\LibUSB\AU6485\AU6485FT.exe", "", "", SW_SHOW)
    End If
    
    SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags
    

End Sub

Public Sub AU6485_SetNBMD()
Dim rt2
Dim EntryTime As Long
Dim PassingTime As Long
Dim mMsg As msg
    
    winHwnd = FindWindow(vbNullString, "AU6485FT_Tool")
    
    EntryTime = Timer
    
    If winHwnd <> 0 Then
    
        rt2 = PostMessage(winHwnd, WM_FT_SetNBMD, 0&, 0&)
        
        Do
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
    
            PassTime = Timer - EntryTime

        Loop Until AlcorMPMessage = WM_FT_TEST_READY _
                Or PassTime > 2
    
    End If
    
End Sub

Public Sub AU6485_SetNormalMD()
Dim rt2
Dim EntryTime As Long
Dim PassingTime As Long
Dim mMsg As msg
    
    winHwnd = FindWindow(vbNullString, "AU6485FT_Tool")
    
    EntryTime = Timer
    
    If winHwnd <> 0 Then
    
        rt2 = PostMessage(winHwnd, WM_FT_SetNorMD, 0&, 0&)
        
        Do
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
    
            PassTime = Timer - EntryTime

        Loop Until AlcorMPMessage = WM_FT_TEST_READY _
                Or PassTime > 2
    
    End If
    
End Sub


Public Function AU6485_SD_Test() As Byte
Dim rt2
Dim EntryTime As Long
Dim PassingTime As Long
Dim mMsg As msg
    
    AU6485_SD_Test = 0
    
    winHwnd = FindWindow(vbNullString, "AU6485FT_Tool")
    
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
                Or AlcorMPMessage = WM_FT_SPEED_ERROR _
                Or PassTime > 8
    
    End If
    
    If AlcorMPMessage = WM_FT_SPEED_ERROR Then
        UsbSpeedTestResult = 2
    ElseIf AlcorMPMessage = WM_FT_SD_PASS Then
        AU6485_SD_Test = 1
    End If

End Function

Public Function AU6485_MS_Test() As Byte
Dim rt2
Dim EntryTime As Long
Dim PassingTime As Long
Dim mMsg As msg
    
    AU6485_MS_Test = 0
    
    winHwnd = FindWindow(vbNullString, "AU6485FT_Tool")
    
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
                Or AlcorMPMessage = WM_FT_SPEED_ERROR _
                Or PassTime > 4
    
    End If
    
    If AlcorMPMessage = WM_FT_SPEED_ERROR Then
        UsbSpeedTestResult = 2
    ElseIf AlcorMPMessage = WM_FT_MS_PASS Then
        AU6485_MS_Test = 1
    End If

End Function

Public Sub Clear_Btn_Click_AU6485()
Dim rt2
Dim EntryTime As Long
Dim PassingTime As Long
Dim mMsg As msg
    
    winHwnd = FindWindow(vbNullString, "AU6485FT_Tool")
    
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


Public Sub CloseAU6485FT_Tool()
Dim rt2 As Long
Dim EntryTime As Long
Dim PassingTime As Long
Dim mMsg As msg
Dim HangHandle As Long

    winHwnd = FindWindow(vbNullString, "AU6485FT_Tool")
    
    EntryTime = Timer
    
    If winHwnd <> 0 Then
        Do
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
    
            PassingTime = Timer - EntryTime

        Loop Until AlcorMPMessage = WM_FT_TEST_READY _
                Or PassingTime > 1
        
        
        Do
'            rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
'            Call MsecDelay(0.5)
            KillProcess ("AU6485FT.exe")
            
            HangHandle = FindWindow(vbNullString, "AU6485FT_Tool: AU6485FT.exe - 應用程式錯誤")
            If HangHandle <> 0 Then
                rt2 = PostMessage(HangHandle, WM_CLOSE, 0&, 0&)
            End If
            
            Call MsecDelay(0.2)
            winHwnd = FindWindow(vbNullString, "AU6485FT_Tool")
        Loop While winHwnd <> 0
    End If
    
End Sub

Public Sub AU6485FTTestSub()

    If (ChipName = "AU6485AFF25") Or (ChipName = "AU6485BFF25") Or (ChipName = "AU6485HFF25") Then
        Call AU6485AFF25TestSub 'Normal Mode
    ElseIf (ChipName = "AU6485CFF25") Or (ChipName = "AU6485DFF25") Or _
           (ChipName = "AU6485IFF25") Or (ChipName = "AU6485JFF25") Then
        Call AU6485CFF25TestSub 'NBMD
    ElseIf (ChipName = "AU6485LFF25") Then
        Call AU6485LFF25TestSub 'NBMD + LED
    ElseIf (ChipName = "AU6485AFF05") Or (ChipName = "AU6485HFF05") Then
        Call AU6485AFF05TestSub 'Normal Mode FT3
    ElseIf (ChipName = "AU6485CFF05") Then
        Call AU6485CFF05TestSub 'NBMD Mode FT3
    ElseIf (ChipName = "AU6485AFS15") Or (ChipName = "AU6485HFS15") Then
        Call AU6485AFS15TestSub 'Normal Mode ST1 (3.6V)
    ElseIf (ChipName = "AU6485CFS15") Then
        Call AU6485CFS15TestSub 'NBMD Mode ST1 (3.6V)
    End If


End Sub

Public Sub AU6485LFF24TestSub()
 
Dim OldTimer
Dim PassTime
Dim rt2
Dim mMsg As msg
Dim TempRes As Integer
Dim TempStr As String
Dim LEDVal As Long
Dim k As Integer


'AU6465-GCF_GBF 28QFN SOCKET
'==============================     ==============================
'P1A   8 7 6 5  4 3 2 1             P1B   8 7 6 5  4 3 2 1
'      | | | |  | | | |                   | | | |  | | | |
'                 M S E                   L
'                 S D N                   E
'                     A                   D
'
'==============================     ==============================


rv0 = 0 'Enum
rv1 = 0 'SD
rv2 = 0 'MS
rv3 = 0 'NBMD
rv4 = 0 'LED

If PCI7248InitFinish = 0 Then
    PCI7248Exist
End If

CardResult = DO_WritePort(card, Channel_P1A, &HFC)
Call MsecDelay(0.2)

AlcorMPMessage = 0


'===================== Wait AP Ready =====================

winHwnd = FindWindow(vbNullString, "AU6485FT_Tool")
If winHwnd = 0 Then

    Call LoadAP_Click_AU6485
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
        TestResult = "Bin2"
        Tester.Label2.Caption = "Bin2"
        Tester.Label2.BackColor = RGB(255, 0, 0)
        Tester.Label9.BackColor = RGB(255, 0, 0)
        Call CloseAU6485FT_Tool
        CardResult = DO_WritePort(card, Channel_P1A, &HFF)
        Call MsecDelay(0.2)
        Exit Sub
    End If
    
    Call AU6485_SetNBMD

End If

Call Clear_Btn_Click_AU6485

'===================== Find Device =====================

rv0 = WaitDevOn("pid_6366")
Call NewLabelMenu(0, "UnknowDevice", rv0, 1)

If rv0 <> 1 Then
    GoTo AU6485ResultLabel
End If

'===================== Check LED =====================
CardResult = DI_ReadPort(card, Channel_P1B, LEDVal)
If rv0 = 1 Then
    For k = 0 To 2
        If (LEDVal And &H1) = &H0 Then          'LED ON
            rv4 = 1
            Exit For
        End If
        Call MsecDelay(0.1)
    Next
End If

'===================== Connect SD & Test =====================
'CardResult = DO_WritePort(card, Channel_P1A, &HFC)
'Call MsecDelay(0.2)

rv1 = AU6485_SD_Test()
Call NewLabelMenu(1, "SD Card", rv1, rv0)
If rv1 <> 1 Then
    GoTo AU6485ResultLabel
End If

'===================== Connect MS & Test =====================
CardResult = DO_WritePort(card, Channel_P1A, &HFA)
Call MsecDelay(0.2)

rv2 = AU6485_MS_Test()
Call NewLabelMenu(2, "MS Card", rv2, rv1)
If rv2 <> 1 Then
    GoTo AU6485ResultLabel
End If


'===================== NBMD Test =====================
If rv2 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &HFE)
    Call MsecDelay(0.3)
    
    If GetDeviceName_NoReply("vid_058f") = "" Then
        rv3 = 1
    End If
    
    Call NewLabelMenu(3, "NB Mode", rv3, rv2)
End If

If (rv3 = 1) And (rv4 = 1) Then
    If rv4 = 1 Then
        rv4 = 0
        For k = 0 To 2
            CardResult = DI_ReadPort(card, Channel_P1B, LEDVal)
            If (LEDVal And &H1) = &H1 Then      'LED OFF
                rv4 = 1
                Exit For
            End If
            Call MsecDelay(0.1)
        Next
    End If
    
    Call NewLabelMenu(4, "LED On/Off", rv4, rv3)

End If


AU6485ResultLabel:
    
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
    WaitDevOFF ("pid_6366")
    
    
    If (rv0 <> 1) Then
        TestResult = "Bin2"     'Unknow
    
    ElseIf (UsbSpeedTestResult = 2) Then    'Speed error
        TestResult = "Bin4"
    
    ElseIf (rv1 <> 1) Then
        TestResult = "Bin3"     'SD Fail
        
    ElseIf (rv2 <> 1) Then
        TestResult = "Bin5"     'MS Fail
        
    ElseIf (rv3 <> 1) Then
        TestResult = "Bin4"     'NBMD Fail
    
    ElseIf (rv4 <> 1) Then
        TestResult = "Bin4"     'LED Fail
        
    ElseIf (rv0 * rv1 * rv2 * rv3 * rv4 = 1) Then
        TestResult = "PASS"
        
    End If
    
    If TestResult <> "PASS" Then
        Call CloseAU6485FT_Tool
    End If
    

Call MsecDelay(0.2)

End Sub
Public Sub AU6485LFF25TestSub()

'2013/10/11
'PM request FT2 test flow using 3.5V(V33) input.

Dim OldTimer
Dim PassTime
Dim rt2
Dim mMsg As msg
Dim TempRes As Integer
Dim TempStr As String
Dim LEDVal As Long
Dim k As Integer


'AU6465-GCF_GBF 28QFN SOCKET
'==============================     ==============================
'P1A   8 7 6 5  4 3 2 1             P1B   8 7 6 5  4 3 2 1
'      | | | |  | | | |                   | | | |  | | | |
'                 M S E                   L
'                 S D N                   E
'                     A                   D
'
'==============================     ==============================


rv0 = 0 'Enum
rv1 = 0 'SD
rv2 = 0 'MS
rv3 = 0 'NBMD
rv4 = 0 'LED

If PCI7248InitFinish_Sync = 0 Then
    PCI7248Exist_P1C_Sync
End If

Call PowerSet2(0, "3.6", "0.5", 1, "3.6", "0.5", 1)
SetSiteStatus (RunHV)
CardResult = DO_WritePort(card, Channel_P1A, &HFC)
Call MsecDelay(0.2)

AlcorMPMessage = 0


'===================== Wait AP Ready =====================

winHwnd = FindWindow(vbNullString, "AU6485FT_Tool")
If winHwnd = 0 Then

    Call LoadAP_Click_AU6485
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
        TestResult = "Bin2"
        Tester.Label2.Caption = "Bin2"
        Tester.Label2.BackColor = RGB(255, 0, 0)
        Tester.Label9.BackColor = RGB(255, 0, 0)
        Call CloseAU6485FT_Tool
        CardResult = DO_WritePort(card, Channel_P1A, &HFF)
        Call MsecDelay(0.2)
        Exit Sub
    End If
    
    Call AU6485_SetNBMD
    
End If

Call Clear_Btn_Click_AU6485

'===================== Find Device =====================

rv0 = WaitDevOn("pid_6366")
Call NewLabelMenu(0, "UnknowDevice", rv0, 1)

If rv0 <> 1 Then
    GoTo AU6485ResultLabel
End If

'===================== Check LED =====================
CardResult = DI_ReadPort(card, Channel_P1B, LEDVal)
If rv0 = 1 Then
    For k = 0 To 2
        If (LEDVal And &H1) = &H0 Then          'LED ON
            rv4 = 1
            Exit For
        End If
        Call MsecDelay(0.1)
    Next
End If

'===================== Connect SD & Test =====================
'CardResult = DO_WritePort(card, Channel_P1A, &HFC)
'Call MsecDelay(0.1)

rv1 = AU6485_SD_Test()
Call NewLabelMenu(1, "SD Card", rv1, rv0)
If rv1 <> 1 Then
    GoTo AU6485ResultLabel
End If


'===================== Connect MS & Test =====================
CardResult = DO_WritePort(card, Channel_P1A, &HFA)
Call MsecDelay(0.1)

rv2 = AU6485_MS_Test()
Call NewLabelMenu(2, "MS Card", rv2, rv1)
If rv2 <> 1 Then
    GoTo AU6485ResultLabel
End If

'===================== NBMD Test =====================
If rv2 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &HFE)
    Call MsecDelay(0.3)
    
    If GetDeviceName_NoReply("vid_058f") = "" Then
        rv3 = 1
    End If
    
    Call NewLabelMenu(3, "NB Mode", rv3, rv2)
End If

    
If (rv3 = 1) And (rv4 = 1) Then
    If rv4 = 1 Then
        rv4 = 0
        For k = 0 To 2
            CardResult = DI_ReadPort(card, Channel_P1B, LEDVal)
            If (LEDVal And &H1) = &H1 Then      'LED OFF
                rv4 = 1
                Exit For
            End If
            Call MsecDelay(0.1)
        Next
    End If
    
    Call NewLabelMenu(4, "LED On/Off", rv4, rv3)

End If


AU6485ResultLabel:
    
    SetSiteStatus (HVDone)
    Call WaitAnotherSiteDone(HVDone, 3#)
    Call PowerSet2(0, "0.0", "0.5", 1, "0.0", "0.5", 1)
    SetSiteStatus (SiteUnknow)
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
    WaitDevOFF ("pid_6366")
    
    
    If (rv0 <> 1) Then
        TestResult = "Bin2"     'Unknow
    
    ElseIf (UsbSpeedTestResult = 2) Then    'Speed error
        TestResult = "Bin4"
    
    ElseIf (rv1 <> 1) Then
        TestResult = "Bin3"     'SD Fail
        
    ElseIf (rv2 <> 1) Then
        TestResult = "Bin5"     'MS Fail
        
    ElseIf (rv3 <> 1) Then
        TestResult = "Bin4"     'NBMD Fail
    
    ElseIf (rv4 <> 1) Then
        TestResult = "Bin4"     'LED Fail
        
    ElseIf (rv0 * rv1 * rv2 * rv3 * rv4 = 1) Then
        TestResult = "PASS"
        
    End If
    
    If TestResult <> "PASS" Then
        Call CloseAU6485FT_Tool
    End If
    

Call MsecDelay(0.2)

End Sub

Public Sub AU6485AFS15TestSub()
 
'2013/10/11: 04 => 05 test flow no changed
'just update flow name with FT2
  
Dim OldTimer
Dim PassTime
Dim rt2
Dim mMsg As msg
Dim TempRes As Integer
Dim TempStr As String


'AU6465-GCF_GBF 28QFN SOCKET
'==============================
'P1A   8 7 6 5  4 3 2 1
'      | | | |  | | | |
'                 M S E
'                 S D N
'                     A
'
'==============================


rv0 = 0 'Enum
rv1 = 0 'SD
rv2 = 0 'MS
'rv3 = 0 'NBMD

If PCI7248InitFinish_Sync = 0 Then
    PCI7248Exist_P1C_Sync
End If

Call PowerSet2(0, "3.6", "0.5", 1, "3.6", "0.5", 1)
SetSiteStatus (RunHV)
CardResult = DO_WritePort(card, Channel_P1A, &HFE)
Call MsecDelay(0.2)

AlcorMPMessage = 0


'===================== Wait AP Ready =====================

winHwnd = FindWindow(vbNullString, "AU6485FT_Tool")
If winHwnd = 0 Then

    Call LoadAP_Click_AU6485
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
        TestResult = "Bin2"
        Tester.Label2.Caption = "Bin2"
        Tester.Label2.BackColor = RGB(255, 0, 0)
        Tester.Label9.BackColor = RGB(255, 0, 0)
        Call CloseAU6485FT_Tool
        CardResult = DO_WritePort(card, Channel_P1A, &HFF)
        Call MsecDelay(0.2)
        Exit Sub
    End If
    
    Call AU6485_SetNormalMD
    
End If

Call Clear_Btn_Click_AU6485

'===================== Find Device =====================

rv0 = WaitDevOn("pid_6366")
Call NewLabelMenu(0, "UnknowDevice", rv0, 1)

If rv0 <> 1 Then
    GoTo AU6485ResultLabel
End If


'===================== Connect SD & Test =====================
CardResult = DO_WritePort(card, Channel_P1A, &HFC)
Call MsecDelay(0.1)

rv1 = AU6485_SD_Test()
Call NewLabelMenu(1, "SD Card", rv1, rv0)
If rv1 <> 1 Then
    GoTo AU6485ResultLabel
End If


'===================== Connect MS & Test =====================
CardResult = DO_WritePort(card, Channel_P1A, &HFA)
Call MsecDelay(0.1)

rv2 = AU6485_MS_Test()
Call NewLabelMenu(2, "MS Card", rv2, rv1)
If rv2 <> 1 Then
    GoTo AU6485ResultLabel
End If


AU6485ResultLabel:
    
    
    SetSiteStatus (HVDone)
    Call WaitAnotherSiteDone(HVDone, 3#)
    Call PowerSet2(0, "0.0", "0.5", 1, "0.0", "0.5", 1)
    SetSiteStatus (SiteUnknow)
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
    WaitDevOFF ("pid_6366")
    
    
    If (rv0 <> 1) Then
        TestResult = "Bin2"     'Unknow
    
    ElseIf (UsbSpeedTestResult = 2) Then    'Speed error
        TestResult = "Bin4"
    
    ElseIf (rv1 <> 1) Then
        TestResult = "Bin3"     'SD Fail
        
    ElseIf (rv2 <> 1) Then
        TestResult = "Bin5"     'MS Fail
        
        
    ElseIf (rv0 * rv1 * rv2 = 1) Then
        TestResult = "PASS"
        
    End If
    
    If TestResult <> "PASS" Then
        Call CloseAU6485FT_Tool
    End If
    

Call MsecDelay(0.3)

End Sub

Public Sub AU6485AFF25TestSub()

'2013/10/11
'PM request FT2 test flow using 3.5V(V33) input.

Dim OldTimer
Dim PassTime
Dim rt2
Dim mMsg As msg
Dim TempRes As Integer
Dim TempStr As String


'AU6465-GCF_GBF 28QFN SOCKET
'==============================
'P1A   8 7 6 5  4 3 2 1
'      | | | |  | | | |
'                 M S E
'                 S D N
'                     A
'
'==============================


rv0 = 0 'Enum
rv1 = 0 'SD
rv2 = 0 'MS
'rv3 = 0 'NBMD

If PCI7248InitFinish_Sync = 0 Then
    PCI7248Exist_P1C_Sync
End If

Call PowerSet2(0, "3.5", "0.5", 1, "3.5", "0.5", 1)
SetSiteStatus (RunHV)
CardResult = DO_WritePort(card, Channel_P1A, &HFE)
Call MsecDelay(0.2)

AlcorMPMessage = 0


'===================== Wait AP Ready =====================

winHwnd = FindWindow(vbNullString, "AU6485FT_Tool")
If winHwnd = 0 Then

    Call LoadAP_Click_AU6485
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
        TestResult = "Bin2"
        Tester.Label2.Caption = "Bin2"
        Tester.Label2.BackColor = RGB(255, 0, 0)
        Tester.Label9.BackColor = RGB(255, 0, 0)
        Call CloseAU6485FT_Tool
        CardResult = DO_WritePort(card, Channel_P1A, &HFF)
        Call MsecDelay(0.2)
        Exit Sub
    End If
    
    Call AU6485_SetNormalMD
    
End If

Call Clear_Btn_Click_AU6485

'===================== Find Device =====================

rv0 = WaitDevOn("pid_6366")
Call NewLabelMenu(0, "UnknowDevice", rv0, 1)

If rv0 <> 1 Then
    GoTo AU6485ResultLabel
End If


'===================== Connect SD & Test =====================
CardResult = DO_WritePort(card, Channel_P1A, &HFC)
Call MsecDelay(0.1)

rv1 = AU6485_SD_Test()
Call NewLabelMenu(1, "SD Card", rv1, rv0)
If rv1 <> 1 Then
    GoTo AU6485ResultLabel
End If


'===================== Connect MS & Test =====================
CardResult = DO_WritePort(card, Channel_P1A, &HFA)
Call MsecDelay(0.1)

rv2 = AU6485_MS_Test()
Call NewLabelMenu(2, "MS Card", rv2, rv1)
If rv2 <> 1 Then
    GoTo AU6485ResultLabel
End If


AU6485ResultLabel:
    
    
    SetSiteStatus (HVDone)
    Call WaitAnotherSiteDone(HVDone, 3#)
    Call PowerSet2(0, "0.0", "0.5", 1, "0.0", "0.5", 1)
    SetSiteStatus (SiteUnknow)
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
    WaitDevOFF ("pid_6366")
    
    
    If (rv0 <> 1) Then
        TestResult = "Bin2"     'Unknow
    
    ElseIf (UsbSpeedTestResult = 2) Then    'Speed error
        TestResult = "Bin4"
    
    ElseIf (rv1 <> 1) Then
        TestResult = "Bin3"     'SD Fail
        
    ElseIf (rv2 <> 1) Then
        TestResult = "Bin5"     'MS Fail
        
        
    ElseIf (rv0 * rv1 * rv2 = 1) Then
        TestResult = "PASS"
        
    End If
    
    If TestResult <> "PASS" Then
        Call CloseAU6485FT_Tool
    End If
    

Call MsecDelay(0.3)

End Sub

Public Sub AU6485CFS15TestSub()
 
'2013/10/11: 04 => 05 test flow no changed
'just update flow name with FT2
  
Dim OldTimer
Dim PassTime
Dim rt2
Dim mMsg As msg
Dim TempRes As Integer
Dim TempStr As String


'AU6465-GCF_GBF 28QFN SOCKET
'==============================
'P1A   8 7 6 5  4 3 2 1
'      | | | |  | | | |
'                 M S E
'                 S D N
'                     A
'
'==============================


rv0 = 0 'Enum
rv1 = 0 'SD
rv2 = 0 'MS
rv3 = 0 'NBMD

If PCI7248InitFinish_Sync = 0 Then
    PCI7248Exist_P1C_Sync
End If

Call PowerSet2(0, "3.6", "0.5", 1, "3.6", "0.5", 1)
SetSiteStatus (RunHV)
CardResult = DO_WritePort(card, Channel_P1A, &HFC)
Call MsecDelay(0.2)

AlcorMPMessage = 0


'===================== Wait AP Ready =====================

winHwnd = FindWindow(vbNullString, "AU6485FT_Tool")
If winHwnd = 0 Then

    Call LoadAP_Click_AU6485
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
        TestResult = "Bin2"
        Tester.Label2.Caption = "Bin2"
        Tester.Label2.BackColor = RGB(255, 0, 0)
        Tester.Label9.BackColor = RGB(255, 0, 0)
        Call CloseAU6485FT_Tool
        CardResult = DO_WritePort(card, Channel_P1A, &HFF)
        Call MsecDelay(0.2)
        Exit Sub
    End If
    
    Call AU6485_SetNBMD
    
End If

Call Clear_Btn_Click_AU6485

'===================== Find Device =====================

rv0 = WaitDevOn("pid_6366")
Call NewLabelMenu(0, "UnknowDevice", rv0, 1)

If rv0 <> 1 Then
    GoTo AU6485ResultLabel
End If


'===================== Connect SD & Test =====================
'CardResult = DO_WritePort(card, Channel_P1A, &HFC)
'Call MsecDelay(0.1)

rv1 = AU6485_SD_Test()
Call NewLabelMenu(1, "SD Card", rv1, rv0)
If rv1 <> 1 Then
    GoTo AU6485ResultLabel
End If


'===================== Connect MS & Test =====================
CardResult = DO_WritePort(card, Channel_P1A, &HFA)
Call MsecDelay(0.1)

rv2 = AU6485_MS_Test()
Call NewLabelMenu(2, "MS Card", rv2, rv1)
If rv2 <> 1 Then
    GoTo AU6485ResultLabel
End If

'===================== NBMD Test =====================
If rv2 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &HFE)
    Call MsecDelay(0.3)
    
    If GetDeviceName_NoReply("vid_058f") = "" Then
        rv3 = 1
    End If
    
    Call NewLabelMenu(3, "NB Mode", rv3, rv2)
End If


AU6485ResultLabel:
    
    
    SetSiteStatus (HVDone)
    Call WaitAnotherSiteDone(HVDone, 3#)
    Call PowerSet2(0, "0.0", "0.5", 1, "0.0", "0.5", 1)
    SetSiteStatus (SiteUnknow)
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
    WaitDevOFF ("pid_6366")
    
    
    If (rv0 <> 1) Then
        TestResult = "Bin2"     'Unknow
    
    ElseIf (UsbSpeedTestResult = 2) Then    'Speed error
        TestResult = "Bin4"
    
    ElseIf (rv1 <> 1) Then
        TestResult = "Bin3"     'SD Fail
        
    ElseIf (rv2 <> 1) Then
        TestResult = "Bin5"     'MS Fail
        
    ElseIf (rv3 <> 1) Then      'NBMD Fail
        TestResult = "Bin4"
        
    ElseIf (rv0 * rv1 * rv2 * rv3 = 1) Then
        TestResult = "PASS"
        
    End If
    
    If TestResult <> "PASS" Then
        Call CloseAU6485FT_Tool
    End If
    

Call MsecDelay(0.3)

End Sub

Public Sub AU6485CFF25TestSub()
 
'2013/10/11
'PM request FT2 test flow using 3.5V(V33) input.

Dim OldTimer
Dim PassTime
Dim rt2
Dim mMsg As msg
Dim TempRes As Integer
Dim TempStr As String


'AU6465-GCF_GBF 28QFN SOCKET
'==============================
'P1A   8 7 6 5  4 3 2 1
'      | | | |  | | | |
'                 M S E
'                 S D N
'                     A
'
'==============================


rv0 = 0 'Enum
rv1 = 0 'SD
rv2 = 0 'MS
rv3 = 0 'NBMD

If PCI7248InitFinish_Sync = 0 Then
    PCI7248Exist_P1C_Sync
End If

Call PowerSet2(0, "3.5", "0.5", 1, "3.5", "0.5", 1)
SetSiteStatus (RunHV)
CardResult = DO_WritePort(card, Channel_P1A, &HFC)
Call MsecDelay(0.2)

AlcorMPMessage = 0


'===================== Wait AP Ready =====================

winHwnd = FindWindow(vbNullString, "AU6485FT_Tool")
If winHwnd = 0 Then

    Call LoadAP_Click_AU6485
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
        TestResult = "Bin2"
        Tester.Label2.Caption = "Bin2"
        Tester.Label2.BackColor = RGB(255, 0, 0)
        Tester.Label9.BackColor = RGB(255, 0, 0)
        Call CloseAU6485FT_Tool
        CardResult = DO_WritePort(card, Channel_P1A, &HFF)
        Call MsecDelay(0.2)
        Exit Sub
    End If
    
    Call AU6485_SetNBMD
    
End If

Call Clear_Btn_Click_AU6485

'===================== Find Device =====================

rv0 = WaitDevOn("pid_6366")
Call NewLabelMenu(0, "UnknowDevice", rv0, 1)

If rv0 <> 1 Then
    GoTo AU6485ResultLabel
End If


'===================== Connect SD & Test =====================
'CardResult = DO_WritePort(card, Channel_P1A, &HFC)
'Call MsecDelay(0.1)

rv1 = AU6485_SD_Test()
Call NewLabelMenu(1, "SD Card", rv1, rv0)
If rv1 <> 1 Then
    GoTo AU6485ResultLabel
End If


'===================== Connect MS & Test =====================
CardResult = DO_WritePort(card, Channel_P1A, &HFA)
Call MsecDelay(0.1)

rv2 = AU6485_MS_Test()
Call NewLabelMenu(2, "MS Card", rv2, rv1)
If rv2 <> 1 Then
    GoTo AU6485ResultLabel
End If

'===================== NBMD Test =====================
If rv2 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &HFE)
    Call MsecDelay(0.3)
    
    If GetDeviceName_NoReply("vid_058f") = "" Then
        rv3 = 1
    End If
    
    Call NewLabelMenu(3, "NB Mode", rv3, rv2)
End If


AU6485ResultLabel:
    
    
    SetSiteStatus (HVDone)
    Call WaitAnotherSiteDone(HVDone, 3#)
    Call PowerSet2(0, "0.0", "0.5", 1, "0.0", "0.5", 1)
    SetSiteStatus (SiteUnknow)
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
    WaitDevOFF ("pid_6366")
    
    
    If (rv0 <> 1) Then
        TestResult = "Bin2"     'Unknow
    
    ElseIf (UsbSpeedTestResult = 2) Then    'Speed error
        TestResult = "Bin4"
    
    ElseIf (rv1 <> 1) Then
        TestResult = "Bin3"     'SD Fail
        
    ElseIf (rv2 <> 1) Then
        TestResult = "Bin5"     'MS Fail
        
    ElseIf (rv3 <> 1) Then      'NBMD Fail
        TestResult = "Bin4"
        
    ElseIf (rv0 * rv1 * rv2 * rv3 = 1) Then
        TestResult = "PASS"
        
    End If
    
    If TestResult <> "PASS" Then
        Call CloseAU6485FT_Tool
    End If
    

Call MsecDelay(0.3)

End Sub

Public Sub AU6485AFF05TestSub()

'2013/10/11: 04 => 05 test flow no changed
'just update flow name with FT2
 
Dim OldTimer
Dim PassTime
Dim rt2
Dim mMsg As msg
Dim TempRes As Integer
Dim TempStr As String
Dim HV_Done_Flag As Boolean
Dim HV_Result As String
Dim LV_Result As String


If PCI7248InitFinish_Sync = 0 Then
    PCI7248Exist_P1C_Sync
End If


Routine_Label:


If Not HV_Done_Flag Then
    Call PowerSet2(0, "3.6", "0.5", 1, "3.6", "0.5", 1)
    Call MsecDelay(0.3)
    Tester.Print "AU6485 : HV Begin Test ..."
    SetSiteStatus (RunHV)
Else
    Call PowerSet2(0, "3.0", "0.5", 1, "3.0", "0.5", 1)
    Call MsecDelay(0.4)
    Tester.Print vbCrLf & "AU6485 : LV Begin Test ..."
    SetSiteStatus (RunLV)
End If



'AU6465-GCF_GBF 28QFN SOCKET
'==============================
'P1A   8 7 6 5  4 3 2 1
'      | | | |  | | | |
'                 M S E
'                 S D N
'                     A
'
'==============================


rv0 = 0 'Enum
rv1 = 0 'SD
rv2 = 0 'MS
'rv3 = 0 'NBMD


CardResult = DO_WritePort(card, Channel_P1A, &HFE)
Call MsecDelay(0.2)

AlcorMPMessage = 0


'===================== Wait AP Ready =====================

winHwnd = FindWindow(vbNullString, "AU6485FT_Tool")
If winHwnd = 0 Then

    Call LoadAP_Click_AU6485
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
        TestResult = "Bin2"
        Tester.Label2.Caption = "Bin2"
        Tester.Label2.BackColor = RGB(255, 0, 0)
        Tester.Label9.BackColor = RGB(255, 0, 0)
        Call CloseAU6485FT_Tool
        CardResult = DO_WritePort(card, Channel_P1A, &HFF)
        Call MsecDelay(0.2)
        Exit Sub
    End If
    
    Call AU6485_SetNormalMD
    
End If

Call Clear_Btn_Click_AU6485

'===================== Find Device =====================

rv0 = WaitDevOn("pid_6366")
Call NewLabelMenu(0, "UnknowDevice", rv0, 1)

If rv0 <> 1 Then
    GoTo AU6485ResultLabel
End If


'===================== Connect SD & Test =====================
CardResult = DO_WritePort(card, Channel_P1A, &HFC)
Call MsecDelay(0.2)

rv1 = AU6485_SD_Test()
Call NewLabelMenu(1, "SD Card", rv1, rv0)
If rv1 <> 1 Then
    GoTo AU6485ResultLabel
End If


'===================== Connect MS & Test =====================
CardResult = DO_WritePort(card, Channel_P1A, &HFA)
Call MsecDelay(0.2)

rv2 = AU6485_MS_Test()
Call NewLabelMenu(2, "MS Card", rv2, rv1)
If rv2 <> 1 Then
    GoTo AU6485ResultLabel
End If


AU6485ResultLabel:
    
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
    If Not HV_Done_Flag Then
        SetSiteStatus (HVDone)
        Call WaitAnotherSiteDone(HVDone, 3#)
    Else
        SetSiteStatus (LVDone)
        Call WaitAnotherSiteDone(LVDone, 3#)
    End If
    Call PowerSet2(0, "0.0", "0.5", 1, "0.0", "0.5", 1)
    Call MsecDelay(0.2)
    WaitDevOFF ("pid_6366")
    SetSiteStatus (SiteUnknow)
    
    If HV_Done_Flag = False Then
        If rv0 <> 1 Then
            HV_Result = "Bin2"
            Tester.Print "HV Unknow"
        ElseIf rv0 * rv1 * rv2 <> 1 Then
            HV_Result = "Fail"
            Tester.Print "HV Fail"
        ElseIf rv0 * rv1 * rv2 = 1 Then
            HV_Result = "PASS"
            Tester.Print "HV PASS"
        End If
        
        HV_Done_Flag = True
        Call MsecDelay(0.4)
        GoTo Routine_Label
    Else
        If rv0 <> 1 Then
            LV_Result = "Bin2"
            Tester.Print "LV Unknow"
        ElseIf rv0 * rv1 * rv2 <> 1 Then
            LV_Result = "Fail"
            Tester.Print "LV Fail"
        ElseIf rv0 * rv1 * rv2 = 1 Then
            LV_Result = "PASS"
            Tester.Print "LV PASS"
        End If
        
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
    
'    If (rv0 <> 1) Then
'        TestResult = "Bin2"     'Unknow
'
'    ElseIf (UsbSpeedTestResult = 2) Then    'Speed error
'        TestResult = "Bin4"
'
'    ElseIf (rv1 <> 1) Then
'        TestResult = "Bin3"     'SD Fail
'
'    ElseIf (rv2 <> 1) Then
'        TestResult = "Bin5"     'MS Fail
'
'
'    ElseIf (rv0 * rv1 * rv2 = 1) Then
'        TestResult = "PASS"
'
'    End If
    
    If TestResult <> "PASS" Then
        Call CloseAU6485FT_Tool
    End If
    
End Sub

Public Sub AU6485CFF05TestSub()

'20130910
'This code copy form AU6485AFF04
'Add NB Mode test item
'2013/10/11: 04 => 05 test flow no changed
'just update flow name with FT2
 

Dim OldTimer
Dim PassTime
Dim rt2
Dim mMsg As msg
Dim TempRes As Integer
Dim TempStr As String
Dim HV_Done_Flag As Boolean
Dim HV_Result As String
Dim LV_Result As String


If PCI7248InitFinish_Sync = 0 Then
    PCI7248Exist_P1C_Sync
End If


Routine_Label:


If Not HV_Done_Flag Then
    Call PowerSet2(0, "3.6", "0.5", 1, "3.6", "0.5", 1)
    Call MsecDelay(0.3)
    Tester.Print "AU6485 : HV Begin Test ..."
    SetSiteStatus (RunHV)
Else
    Call PowerSet2(0, "3.0", "0.5", 1, "3.0", "0.5", 1)
    Call MsecDelay(0.4)
    Tester.Print vbCrLf & "AU6485 : LV Begin Test ..."
    SetSiteStatus (RunLV)
End If



'AU6465-GCF_GBF 28QFN SOCKET
'==============================
'P1A   8 7 6 5  4 3 2 1
'      | | | |  | | | |
'                 M S E
'                 S D N
'                     A
'
'==============================


rv0 = 0 'Enum
rv1 = 0 'SD
rv2 = 0 'MS
rv3 = 0 'NBMD


CardResult = DO_WritePort(card, Channel_P1A, &HFC)
Call MsecDelay(0.2)

AlcorMPMessage = 0


'===================== Wait AP Ready =====================

winHwnd = FindWindow(vbNullString, "AU6485FT_Tool")
If winHwnd = 0 Then

    Call LoadAP_Click_AU6485
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
        TestResult = "Bin2"
        Tester.Label2.Caption = "Bin2"
        Tester.Label2.BackColor = RGB(255, 0, 0)
        Tester.Label9.BackColor = RGB(255, 0, 0)
        Call CloseAU6485FT_Tool
        CardResult = DO_WritePort(card, Channel_P1A, &HFF)
        Call MsecDelay(0.2)
        Exit Sub
    End If
    
    Call AU6485_SetNBMD
    
End If

Call Clear_Btn_Click_AU6485

'===================== Find Device =====================

rv0 = WaitDevOn("pid_6366")
Call NewLabelMenu(0, "UnknowDevice", rv0, 1)

If rv0 <> 1 Then
    GoTo AU6485ResultLabel
End If


'===================== Connect SD & Test =====================
'CardResult = DO_WritePort(card, Channel_P1A, &HFC)
'Call MsecDelay(0.2)

rv1 = AU6485_SD_Test()
Call NewLabelMenu(1, "SD Card", rv1, rv0)
If rv1 <> 1 Then
    GoTo AU6485ResultLabel
End If


'===================== Connect MS & Test =====================
CardResult = DO_WritePort(card, Channel_P1A, &HFA)
Call MsecDelay(0.2)

rv2 = AU6485_MS_Test()
Call NewLabelMenu(2, "MS Card", rv2, rv1)
If rv2 <> 1 Then
    GoTo AU6485ResultLabel
End If


'===================== NBMD Test =====================
If rv2 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &HFE)
    Call MsecDelay(0.3)
    
    If GetDeviceName_NoReply("vid_058f") = "" Then
        rv3 = 1
    End If
    
    Call NewLabelMenu(3, "NB Mode", rv3, rv2)
End If


AU6485ResultLabel:
    
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
    If Not HV_Done_Flag Then
        SetSiteStatus (HVDone)
        Call WaitAnotherSiteDone(HVDone, 3#)
    Else
        SetSiteStatus (LVDone)
        Call WaitAnotherSiteDone(LVDone, 3#)
    End If
    Call PowerSet2(0, "0.0", "0.5", 1, "0.0", "0.5", 1)
    Call MsecDelay(0.2)
    'WaitDevOFF ("pid_6366")
    SetSiteStatus (SiteUnknow)
    
    If HV_Done_Flag = False Then
        If rv0 <> 1 Then
            HV_Result = "Bin2"
            Tester.Print "HV Unknow"
        ElseIf rv0 * rv1 * rv2 * rv3 <> 1 Then
            HV_Result = "Fail"
            Tester.Print "HV Fail"
        ElseIf rv0 * rv1 * rv2 * rv3 = 1 Then
            HV_Result = "PASS"
            Tester.Print "HV PASS"
        End If
        HV_Done_Flag = True
        Call MsecDelay(0.4)
        GoTo Routine_Label
    Else
        If rv0 <> 1 Then
            LV_Result = "Bin2"
            Tester.Print "LV Unknow"
        ElseIf rv0 * rv1 * rv2 * rv3 <> 1 Then
            LV_Result = "Fail"
            Tester.Print "LV Fail"
        ElseIf rv0 * rv1 * rv2 * rv3 = 1 Then
            LV_Result = "PASS"
            Tester.Print "LV PASS"
        End If
        
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
    
'    If (rv0 <> 1) Then
'        TestResult = "Bin2"     'Unknow
'
'    ElseIf (UsbSpeedTestResult = 2) Then    'Speed error
'        TestResult = "Bin4"
'
'    ElseIf (rv1 <> 1) Then
'        TestResult = "Bin3"     'SD Fail
'
'    ElseIf (rv2 <> 1) Then
'        TestResult = "Bin5"     'MS Fail
'
'
'    ElseIf (rv0 * rv1 * rv2 = 1) Then
'        TestResult = "PASS"
'
'    End If
    
    If TestResult <> "PASS" Then
        Call CloseAU6485FT_Tool
    End If
    
End Sub

