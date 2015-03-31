Attribute VB_Name = "AU6601"
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Public Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As msg) As Long
Public Declare Function TranslateMessage Lib "user32" (lpMsg As msg) As Long

Public Type msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type


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


Public Const AU6601_UnknowFail = 0
Public Const AU6601_DiskFail = 2
Public Const AU6601_RWFail = 3
Public Const AU6601_CardTypeFail = 4
Public Const AU6601_LEDFail = 5
Public Const AU6601_TestPASS = 1

Public Const WM_USER = &H400
Public Const WM_FT_START = WM_USER + &H100
Public Const WM_FT_UNKNOW_FAIL = WM_USER + &H120
Public Const WM_FT_DISK_FAIL = WM_USER + &H150
Public Const WM_FT_RW_FAIL = WM_USER + &H200
Public Const WM_FT_CARDTYPE_FAIL = WM_USER + &H250
Public Const WM_FT_PASS = WM_USER + &H400
Public Const WM_FT_CLEAR = WM_USER + &H500
Public Const WM_FT_GetDevice = WM_USER + &H600
Public Const WM_FT_TEST_READY = WM_USER + &H800

Public Const AU6601_AP_Name = "AU6601_FT.exe"
Public Const AU6601_AP_Title = "AU6601_FT"


Public Function LoadAP_AU6601_FT() As Byte

Dim PassTime As Long
Dim EntryTime As Long
       
    LoadAP_AU6601_FT = 0
    
    ' find window
    winHwnd = FindWindow(vbNullString, AU6601_AP_Title)
     
    ' run program
    If winHwnd = 0 Then
        'ChDir (App.Path)
        Call ShellExecute(Tester.hwnd, "open", App.Path & "\AU6601\" & AU6601_AP_Name, "", "", SW_SHOW)
    Else
        LoadAP_AU6601_FT = 1
        Exit Function
    End If
    
    EntryTime = Timer
    
    Do
        winHwnd = FindWindow(vbNullString, AU6601_AP_Title)
        If winHwnd <> 0 Then
            LoadAP_AU6601_FT = 1
            Call MsecDelay(0.5)
            Exit Do
        End If
        Call MsecDelay(0.5)
        PassTime = Timer - EntryTime
        
    Loop Until (PassTime >= 3)
    
    SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags
    
    ChDir (App.Path)

End Function

Public Sub Close_AU6601_AP()
Dim rt2 As Long
Dim EntryTime As Long
Dim PassingTime As Long
Dim mMsg As msg

    winHwnd = FindWindow(vbNullString, AU6601_AP_Title)
    
    EntryTime = Timer
    
    If winHwnd <> 0 Then
        Do
            rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
            Call MsecDelay(0.3)
        
            winHwnd = FindWindow(vbNullString, AU6601_AP_Title)
            PassingTime = Timer - EntryTime
        Loop Until (winHwnd <> 0) Or (PassingTime >= 2)
    End If
    
End Sub

Public Sub Clear_AU6601_AP()
Dim rt2 As Long
Dim mMsg As msg

    winHwnd = FindWindow(vbNullString, AU6601_AP_Title)
    
    If winHwnd <> 0 Then
        rt2 = PostMessage(winHwnd, WM_FT_CLEAR, 0&, 0&)
    End If
    
End Sub

Public Function StartTest_AU6601_AP(TimeOut As Single) As Byte
Dim rt2 As Long
Dim EntryTime As Long
Dim PassingTime As Long
Dim mMsg As msg
    
    StartTest_AU6601_AP = 0
    AlcorMPMessage = 0
    
    winHwnd = FindWindow(vbNullString, AU6601_AP_Title)
    
    EntryTime = Timer
    
    If winHwnd <> 0 Then
        Do
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
            
            PassingTime = Timer - EntryTime
            
        Loop Until AlcorMPMessage = WM_FT_GetDevice _
                Or PassingTime > 4
    
    End If
    
    PassingTime = 0
    AlcorMPMessage = 0
    
    If winHwnd <> 0 Then
        Call MsecDelay(0.2)
        rt2 = PostMessage(winHwnd, WM_FT_START, 0, 0&)
    Else
        Exit Function
    End If
    
    EntryTime = Timer
    
    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If

        PassingTime = Timer - EntryTime

    Loop Until AlcorMPMessage = WM_FT_UNKNOW_FAIL _
            Or AlcorMPMessage = WM_FT_DISK_FAIL _
            Or AlcorMPMessage = WM_FT_RW_FAIL _
            Or AlcorMPMessage = WM_FT_CARDTYPE_FAIL _
            Or AlcorMPMessage = WM_FT_PASS _
            Or PassingTime > TimeOut
    
    If AlcorMPMessage = WM_FT_UNKNOW_FAIL Then
        StartTest_AU6601_AP = 0
    ElseIf AlcorMPMessage = WM_FT_DISK_FAIL Then
        StartTest_AU6601_AP = 2
    ElseIf AlcorMPMessage = WM_FT_RW_FAIL Then
        StartTest_AU6601_AP = 3
    ElseIf AlcorMPMessage = WM_FT_CARDTYPE_FAIL Then
        StartTest_AU6601_AP = 4
    ElseIf AlcorMPMessage = WM_FT_PASS Then
        StartTest_AU6601_AP = 1
    Else
        StartTest_AU6601_AP = 0
    End If
    
End Function

Public Sub ReSacanPCI()

Dim ReturnPid As String

    ResetHubReturn = Shell(App.Path & "\devcon rescan", vbNormalFocus)
    WaitProcQuit (ResetHubReturn)

    Call MsecDelay(0.2)

End Sub

Public Sub AU6601FTTestSub()

    If ChipName = "AU6601CFF20" Then
        Call AU6601CFF20TestSub
    ElseIf ChipName = "AU6601CFF00" Then
        Call AU6601CFF00TestSub
    End If
    
End Sub

Public Function WaitProcQuit(pId As Long)

On Error Resume Next
Dim objProcess
Dim Pid_Exist As Boolean

    Do
        Pid_Exist = True
        For Each objProcess In GetObject("winmgmts:\\.\root\cimv2:win32_process").instances_
            'Debug.Print objProcess.Handle; objProcess.Name
             If objProcess.Handle = pId Then
                Pid_Exist = False
                'Debug.Print "Ongo"
                Exit For
             End If
        Next
        
    Loop Until (Pid_Exist)

    Set objProcess = Nothing
    
End Function

Public Sub AU6601CFF20TestSub()
 
Dim OldTimer
Dim PassTime
Dim rt2
Dim mMsg As msg
Dim TempRes As Integer
Dim TempStr As String
Dim LightOn As Long
Dim LightCount As Integer

'AU6601-CF 40QFN SOCKET V2
'==============================   ==============================
'P1A   8 7 6 5  4 3 2 1     P1B   8 7 6 5  4 3 2 1
'      | | | |  | | | |           | | | |  | | | |
'                 M S E                          L
'                 S D N                          E
'                     A                          D
'
'==============================   ==============================


rv0 = 0 'Enum
rv1 = 0 'SD
rv2 = 0 'LED_On
rv3 = 0 'MS


LightOn = &HFF

If PCI7248InitFinish_Sync = 0 Then
    PCI7248Exist_P1C_Sync
End If

'===================== Wait AP Ready =====================
winHwnd = FindWindow(vbNullString, AU6601_AP_Title)
If winHwnd = 0 Then

    Call LoadAP_AU6601_FT
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
        Tester.Label9.Caption = "Load AP Fail"
        Call Close_AU6601_AP
        CardResult = DO_WritePort(card, Channel_P1A, &HFF)
        Call MsecDelay(0.2)
        Exit Sub
    End If

End If

Call Clear_AU6601_AP

Call PowerSet2(1, "3.3", "0.5", 1, "3.3", "0.5", 1)
Call MsecDelay(0.1)

SetSiteStatus (RunHV)
Call WaitAnotherSiteDone(RunHV, 2)
'===================== Connect SD & Test =====================
CardResult = DO_WritePort(card, Channel_P1A, &HFC)
ReSacanPCI

rv1 = StartTest_AU6601_AP(6)

If rv1 <> AU6601_TestPASS Then
    Call MsecDelay(0.9)
    CardResult = DO_WritePort(card, Channel_P1A, &HFE)
    Call MsecDelay(0.4)
    CardResult = DO_WritePort(card, Channel_P1A, &HFC)
    Call Clear_AU6601_AP
    rv1 = StartTest_AU6601_AP(6)
End If


If rv1 = AU6601_UnknowFail Then
    rv0 = 0
    Call NewLabelMenu(0, "Enum", rv0, 1)
Else
    rv0 = 1
    Call NewLabelMenu(0, "Enum", rv0, 1)
End If

If rv1 = AU6601_DiskFail Then
    Call NewLabelMenu(1, "SD Disk", rv1, rv0)
ElseIf rv1 = AU6601_RWFail Then
    Call NewLabelMenu(1, "SD R/W", rv1, rv0)
ElseIf rv1 = AU6601_CardTypeFail Then
    Call NewLabelMenu(1, "SD Card Speed/Width", rv1, rv0)
ElseIf rv1 = AU6601_TestPASS Then
    Call NewLabelMenu(1, "SD Card Speed/Width", rv1, rv0)
End If

If rv1 <> 1 Then
    GoTo AU6601ResultLabel
End If

'===================== Connect MS & Test =====================
Call MsecDelay(0.9)  'if < 0.6 will blue-screen
Call Clear_AU6601_AP
CardResult = DO_WritePort(card, Channel_P1A, &HFE)
Call MsecDelay(0.4)

' LED On Test
For LightCount = 1 To 10
    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
    If (LightOn And 1) = 0 Then
        Exit For
    Else
        Call MsecDelay(0.1)
    End If
Next

If ((LightOn And 1) <> 0) Then
    rv2 = 0
    Call NewLabelMenu(2, "LED ON", rv2, rv1)
    Tester.Print "LED On Fail"
    GoTo AU6601ResultLabel
Else
    rv2 = 1
    Call NewLabelMenu(2, "LED ON", rv2, rv1)
End If

CardResult = DO_WritePort(card, Channel_P1A, &HFA)

rv3 = StartTest_AU6601_AP(6)

If rv3 <> AU6601_TestPASS Then
    Call MsecDelay(0.9)
    CardResult = DO_WritePort(card, Channel_P1A, &HFE)
    Call MsecDelay(0.4)
    CardResult = DO_WritePort(card, Channel_P1A, &HFA)
    Call Clear_AU6601_AP
    rv3 = StartTest_AU6601_AP(6)
End If

If rv3 = AU6601_DiskFail Then
    Call NewLabelMenu(3, "MS Disk", rv3, rv2)
ElseIf rv3 = AU6601_RWFail Then
    Call NewLabelMenu(3, "MS R/W", rv3, rv2)
ElseIf rv3 = AU6601_CardTypeFail Then
    Call NewLabelMenu(3, "MS Card Speed/Width", rv3, rv2)
ElseIf rv3 = AU6601_TestPASS Then
    Call NewLabelMenu(3, "MS Card Speed/Width", rv3, rv2)
End If

'===================== LED OFF Test =====================
For LightCount = 1 To 10
    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
    If (LightOn And 1) = 1 Then
        Exit For
    Else
        Call MsecDelay(0.1)
    End If
Next

If ((LightOn And 1) <> 1) Then
    rv2 = 0
    Tester.Print "LED OFF Fail"
End If
Call NewLabelMenu(4, "LED OFF", rv2, rv3)


AU6601ResultLabel:
    
    SetSiteStatus (HVDone)
    Call WaitAnotherSiteDone(HVDone, 9)
    Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
    Call MsecDelay(0.1)
    ReSacanPCI
    
    SetSiteStatus (SiteUnknow)
    
    If (rv0 <> 1) Then
        TestResult = "Bin2"     'Unknow
    
    ElseIf (rv1 <> 1) Then
        TestResult = "Bin3"     'SD Fail
        
    ElseIf (rv2 <> 1) Then
        TestResult = "Bin4"     'LED On/Off Fail
    
    ElseIf (rv3 <> 1) Then      'MS Fail
        TestResult = "Bin5"
             
    ElseIf (rv0 * rv1 * rv2 * rv3 = 1) Then
        TestResult = "PASS"
        
    End If
    
    If TestResult <> "PASS" Then
        Call Close_AU6601_AP
    End If
    
End Sub

Public Sub AU6601CFF00TestSub()
 
Dim OldTimer
Dim PassTime
Dim rt2
Dim mMsg As msg
Dim TempRes As Integer
Dim TempStr As String
Dim LightOn As Long
Dim LightCount As Integer
Dim HV_Done As Boolean
Dim HV_Result As String
Dim LV_Result As String

'AU6601-CF 40QFN SOCKET V2
'==============================   ==============================
'P1A   8 7 6 5  4 3 2 1     P1B   8 7 6 5  4 3 2 1
'      | | | |  | | | |           | | | |  | | | |
'                 M S E                          L
'                 S D N                          E
'                     A                          D
'
'==============================   ==============================

HV_Done = False
HV_Result = ""
LV_Result = ""


Routine_Label:



Tester.Label3.BackColor = RGB(255, 255, 255)
Tester.Label4.BackColor = RGB(255, 255, 255)
Tester.Label5.BackColor = RGB(255, 255, 255)
Tester.Label6.BackColor = RGB(255, 255, 255)
Tester.Label7.BackColor = RGB(255, 255, 255)
Tester.Label8.BackColor = RGB(255, 255, 255)

rv0 = 0 'Enum
rv1 = 0 'SD
rv2 = 0 'LED_On
rv3 = 0 'MS

LightOn = &HFF

If PCI7248InitFinish_Sync = 0 Then
    PCI7248Exist_P1C_Sync
End If

'===================== Wait AP Ready =====================
winHwnd = FindWindow(vbNullString, AU6601_AP_Title)
If winHwnd = 0 Then

    Call LoadAP_AU6601_FT
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
        Tester.Label9.Caption = "Load AP Fail"
        Call Close_AU6601_AP
        CardResult = DO_WritePort(card, Channel_P1A, &HFF)
        Call MsecDelay(0.2)
        Exit Sub
    End If

End If

Call Clear_AU6601_AP

If Not HV_Done Then
    Call PowerSet2(1, "3.6", "0.5", 1, "3.6", "0.5", 1)
    Call MsecDelay(0.1)
    SetSiteStatus (RunHV)
    Call WaitAnotherSiteDone(RunHV, 2)
    Tester.Print "AU6601CF Begin HV(3.6) Test ..."
Else
    Call PowerSet2(1, "3.0", "0.5", 1, "3.0", "0.5", 1)
    Call MsecDelay(0.1)
    SetSiteStatus (RunLV)
    Call WaitAnotherSiteDone(RunLV, 2)
    Tester.Print "AU6601CF Begin LV(3.0) Test ..."
End If
'===================== Connect SD & Test =====================
CardResult = DO_WritePort(card, Channel_P1A, &HFC)
ReSacanPCI

rv1 = StartTest_AU6601_AP(6)

If rv1 <> AU6601_TestPASS Then
    Call MsecDelay(0.9)
    CardResult = DO_WritePort(card, Channel_P1A, &HFE)
    Call MsecDelay(0.4)
    CardResult = DO_WritePort(card, Channel_P1A, &HFC)
    Call Clear_AU6601_AP
    rv1 = StartTest_AU6601_AP(6)
End If


If rv1 = AU6601_UnknowFail Then
    rv0 = 0
    Call NewLabelMenu(0, "Enum", rv0, 1)
Else
    rv0 = 1
    Call NewLabelMenu(0, "Enum", rv0, 1)
End If

If rv1 = AU6601_DiskFail Then
    Call NewLabelMenu(1, "SD Disk", rv1, rv0)
ElseIf rv1 = AU6601_RWFail Then
    Call NewLabelMenu(1, "SD R/W", rv1, rv0)
ElseIf rv1 = AU6601_CardTypeFail Then
    Call NewLabelMenu(1, "SD Card Speed/Width", rv1, rv0)
ElseIf rv1 = AU6601_TestPASS Then
    Call NewLabelMenu(1, "SD Card Speed/Width", rv1, rv0)
End If

If rv1 <> 1 Then
    GoTo AU6601ResultLabel
End If

'===================== Connect MS & Test =====================
Call MsecDelay(0.9)  'if < 0.6 will blue-screen
Call Clear_AU6601_AP
CardResult = DO_WritePort(card, Channel_P1A, &HFE)
Call MsecDelay(0.4)

' LED On Test
For LightCount = 1 To 10
    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
    If (LightOn And 1) = 0 Then
        Exit For
    Else
        Call MsecDelay(0.1)
    End If
Next

If ((LightOn And 1) <> 0) Then
    rv2 = 0
    Call NewLabelMenu(2, "LED ON", rv2, rv1)
    Tester.Print "LED On Fail"
    GoTo AU6601ResultLabel
Else
    rv2 = 1
    Call NewLabelMenu(2, "LED ON", rv2, rv1)
End If

CardResult = DO_WritePort(card, Channel_P1A, &HFA)

rv3 = StartTest_AU6601_AP(6)

If rv3 <> AU6601_TestPASS Then
    Call MsecDelay(0.9)
    CardResult = DO_WritePort(card, Channel_P1A, &HFE)
    Call MsecDelay(0.4)
    CardResult = DO_WritePort(card, Channel_P1A, &HFA)
    Call Clear_AU6601_AP
    rv3 = StartTest_AU6601_AP(6)
End If

If rv3 = AU6601_DiskFail Then
    Call NewLabelMenu(3, "MS Disk", rv3, rv2)
ElseIf rv3 = AU6601_RWFail Then
    Call NewLabelMenu(3, "MS R/W", rv3, rv2)
ElseIf rv3 = AU6601_CardTypeFail Then
    Call NewLabelMenu(3, "MS Card Speed/Width", rv3, rv2)
ElseIf rv3 = AU6601_TestPASS Then
    Call NewLabelMenu(3, "MS Card Speed/Width", rv3, rv2)
End If

'===================== LED OFF Test =====================
For LightCount = 1 To 10
    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
    If (LightOn And 1) = 1 Then
        Exit For
    Else
        Call MsecDelay(0.1)
    End If
Next

If ((LightOn And 1) <> 1) Then
    rv2 = 0
    Tester.Print "LED OFF Fail"
End If
Call NewLabelMenu(4, "LED OFF", rv2, rv3)


AU6601ResultLabel:
    
    
    If Not HV_Done Then
        SetSiteStatus (HVDone)
        Call WaitAnotherSiteDone(HVDone, 12)
        Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
        CardResult = DO_WritePort(card, Channel_P1A, &HFF)
        Call MsecDelay(0.1)
        ReSacanPCI
        
        
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
        
        SetSiteStatus (SiteUnknow)
        
        HV_Done = True
        
        GoTo Routine_Label
    Else
        SetSiteStatus (LVDone)
        Call WaitAnotherSiteDone(LVDone, 12)
        Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
        CardResult = DO_WritePort(card, Channel_P1A, &HFF)
        Call MsecDelay(0.1)
        ReSacanPCI
        
        SetSiteStatus (SiteUnknow)
        
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
    
    If TestResult <> "PASS" Then
        Call Close_AU6601_AP
    End If
    
End Sub
