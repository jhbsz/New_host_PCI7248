Attribute VB_Name = "AU9540"
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

Public RunReleaseMemCount As Integer

Option Explicit


Public Const WM_USER = &H400
Public Const WM_FT_TEST_READY = WM_USER + &H800
Public Const WM_FT_TEST_HID_START = WM_USER + &H100
Public Const WM_FT_TEST_CCID_START = WM_USER + &H150
Public Const WM_FT_TEST_UNKNOW = WM_USER + &H200
Public Const WM_FT_TEST_ATR_FAIL = WM_USER + &H220
Public Const WM_FT_TEST_APDU_FAIL = WM_USER + &H240
Public Const WM_FT_TEST_PASS = WM_USER + &H400


Public Sub LoadAP_Click_AU9540()
Dim TimePass
Dim rt2
    
    ' find window
    winHwnd = FindWindow(vbNullString, "AU9540 USBHID AP")
     
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(Tester.hwnd, "open", App.Path & "\AU9540\AU9540FT_AP.exe", "", "", SW_SHOW)
    End If
    
    SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags
    

End Sub

Public Sub LoadAP_Click_SmartCard_COM_Mode()
Dim TimePass
Dim rt2
    
    ' find window
    winHwnd = FindWindow(vbNullString, "SmartCard_COM_Mode_Tester")
     
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(Tester.hwnd, "open", App.Path & "\AU9562\SmartCard_COM_Mode_Tester.exe", "", "", SW_SHOW)
    End If
    
    SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags
    

End Sub

Public Sub StartRWTest_Click_AU9540_HID()
Dim rt2
    
    winHwnd = FindWindow(vbNullString, "AU9540 USBHID AP")
    rt2 = PostMessage(winHwnd, WM_FT_TEST_HID_START, 0&, 0&)

End Sub

Public Sub StartRWTest_Click_AU9540_CCID()
Dim rt2
    
    winHwnd = FindWindow(vbNullString, "AU9540 USBHID AP")
    rt2 = PostMessage(winHwnd, WM_FT_TEST_CCID_START, 0&, 0&)

End Sub

Public Sub StartRWTest_Click_SmartCard_COM_Mode()
Dim rt2
    
    winHwnd = FindWindow(vbNullString, "SmartCard_COM_Mode_Tester")
    rt2 = PostMessage(winHwnd, WM_COM_START_TEST, 0&, 0&)

End Sub

Public Sub CloseAU9540AP()
Dim rt2 As Long
Dim mMsg As msg

    winHwnd = FindWindow(vbNullString, "AU9540 USBHID AP")
 
    If winHwnd <> 0 Then
        Do
            rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
            Call MsecDelay(0.5)
            winHwnd = FindWindow(vbNullString, "AU9540 USBHID AP")
        Loop While winHwnd <> 0
    End If
    
End Sub

Public Sub Close_SmartCard_COM_Mode()
Dim rt2 As Long
Dim EntryTime As Long
Dim PassingTime As Long
Dim mMsg As msg
Dim HangHandle As Long

    winHwnd = FindWindow(vbNullString, "SmartCard_COM_Mode_Tester")
 
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

            KillProcess ("SmartCard_COM_Mode_Tester.exe")
            
            HangHandle = FindWindow(vbNullString, "SmartCard_COM_Mode_Tester: SmartCard_COM_Mode_Tester.exe - 應用程式錯誤")
            If HangHandle <> 0 Then
                rt2 = PostMessage(HangHandle, WM_COM_TEST_CLOSE, 0&, 0&)
            End If
            
            Call MsecDelay(0.2)
            winHwnd = FindWindow(vbNullString, "SmartCard_COM_Mode_Tester")
        Loop While winHwnd <> 0
    
    End If
    
End Sub

Public Sub AU9540TestSub()
 
Dim OldTimer
Dim PassTime
Dim rt2
Dim mMsg As msg
Dim TempRes As Integer
Dim TempStr As String
Dim LightOn As Long
Dim LightCount As Integer

If PCI7248InitFinish = 0 Then
    PCI7248Exist
End If

Call MsecDelay(0.5)
CardResult = DO_WritePort(card, Channel_P1A, &HFE)
Call MsecDelay(0.2)

AlcorMPMessage = 0

winHwnd = FindWindow(vbNullString, "AU9540 USBHID AP")

If winHwnd = 0 Then

    Call LoadAP_Click_AU9540
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
    Tester.Print "AP Ready Time="; PassTime

    If PassTime > 5 Then
        TestResult = "Bin3"
        Tester.Label2.Caption = "Bin3"
        Tester.Label2.BackColor = RGB(255, 0, 0)
        Tester.Label9.BackColor = RGB(255, 0, 0)
        Call CloseAU9540AP
        CardResult = DO_WritePort(card, Channel_P1A, &HFF)
        Call MsecDelay(0.2)
        Exit Sub
    End If

End If

TempRes = WaitDevOn("vid_1059")
Call MsecDelay(0.3)

If TempRes <> 1 Then
    TestResult = "Bin2"
    Tester.Label2.Caption = "Bin2"
    Tester.Label2.BackColor = RGB(255, 0, 0)
    Tester.Label9.Caption = "Find Device Fail"
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
    AU6256Unknow = AU6256Unknow + 1
    If AU6256Unknow > 5 Then
        Shell "cmd /c shutdown -r  -t 0", vbHide
    End If
    Call MsecDelay(0.2)
    Exit Sub
End If

Do
    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
    LightCount = LightCount + 1
    Call MsecDelay(0.1)
Loop While (LightOn <> &HFC) And (LightCount < 10)


TempStr = GetDeviceName_NoReply("vid_1059")

Tester.Print "AU9540 begin test........"

If InStr(TempStr, "pid_0018") Then              'HID --- Vid: 1059, Pid: 0018
    Call StartRWTest_Click_AU9540_HID
ElseIf InStr(TempStr, "pid_0017") Then          'CCID -- Vid: 1059, Pid: 0017
    Call StartRWTest_Click_AU9540_CCID
End If

OldTimer = Timer

Do
    If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
        AlcorMPMessage = mMsg.message
        TranslateMessage mMsg
        DispatchMessage mMsg
    End If
         
    PassTime = Timer - OldTimer
       
Loop Until AlcorMPMessage = WM_FT_TEST_UNKNOW _
      Or AlcorMPMessage = WM_FT_TEST_ATR_FAIL _
      Or AlcorMPMessage = WM_FT_TEST_APDU_FAIL _
      Or AlcorMPMessage = WM_FT_TEST_PASS _
      Or PassTime > 5
    
    
Tester.Print "RW work Time="; PassTime
        
'===========================================================
'  RW Time Out Fail
'===========================================================

If PassTime > 5 Then
    TestResult = "Bin3"
    Tester.Label2.Caption = "Bin3"
    Tester.Label9.Caption = "Bin3: Test TimeOut"
    Call CloseAU9540AP
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
    Call MsecDelay(0.2)
    Exit Sub
End If
               
Select Case AlcorMPMessage
            
    Case WM_FT_TEST_UNKNOW
        TestResult = "Bin2"
        Tester.Label2.Caption = "Bin2"
        Tester.Label2.BackColor = RGB(255, 0, 0)
        Tester.Label9.Caption = "Bin2: UnKnow Fail"
        AU6256Unknow = AU6256Unknow + 1
        If AU6256Unknow > 5 Then
            Shell "cmd /c shutdown -r  -t 0", vbHide
        End If
        Call CloseAU9540AP
        
    Case WM_FT_TEST_ATR_FAIL
        TestResult = "Bin4"
        Tester.Label2.Caption = "Bin4"
        Tester.Label2.BackColor = RGB(255, 0, 0)
        Tester.Label9.Caption = "Bin4: Get ATR Fail"
        AU6256Unknow = 0
        Call CloseAU9540AP
        
    Case WM_FT_TEST_APDU_FAIL
        TestResult = "Bin5"
        Tester.Label2.Caption = "Bin5"
        Tester.Label2.BackColor = RGB(255, 0, 0)
        Tester.Label9.Caption = "Bin5: Get APDU Fail"
        AU6256Unknow = 0
        Call CloseAU9540AP
        
    Case WM_FT_TEST_PASS
        If (LightOn <> &HFC) Then
            TestResult = "Bin3"
            Tester.Label2.Caption = "Bin3"
            Tester.Label2.BackColor = RGB(255, 0, 0)
            Tester.Label9.Caption = "Bin3:LED Fail"
            AU6256Unknow = 0
        Else
            TestResult = "PASS"
            Tester.Label2.Caption = "PASS"
            Tester.Label2.BackColor = RGB(0, 255, 0)
            Tester.Label9.Caption = "PASS"
            AU6256Unknow = 0
        End If
    Case Else
        TestResult = "Bin2"
        Tester.Label2.Caption = "Bin2"
        Tester.Label2.BackColor = RGB(255, 0, 0)
        Tester.Label9.Caption = "Bin2:Undefine Fail"

End Select

CardResult = DO_WritePort(card, Channel_P1A, &HFF)
WaitDevOFF ("vid_1059")
Call MsecDelay(0.3)

End Sub

Public Sub AU9540Test_HLV_Sub()
 
Dim OldTimer
Dim PassTime
Dim rt2
Dim mMsg As msg
Dim TempRes As Integer
Dim TempStr As String
Dim LightOn As Long
Dim LightCount As Integer
Dim HV_Done_Flag As Boolean
Dim HV_Result, LV_Result As String

If PCI7248InitFinish_Sync = 0 Then
    PCI7248Exist_P1C_Sync
End If


Routine_Label:


If (Not HV_Done_Flag) Then
    Tester.Print "HV_Test(3.6V) Begin ..."
    Call PowerSet2(0, "3.6", "0.3", 1, "3.6", "0.3", 1)
    SetSiteStatus (RunHV)
Else
    Tester.Print vbCrLf & "LV_Test(3.0V) Begin ..."
    Call PowerSet2(0, "3.0", "0.3", 1, "3.0", "0.3", 1)
    SetSiteStatus (RunLV)
End If


Call MsecDelay(0.2)
CardResult = DO_WritePort(card, Channel_P1A, &HFE)
Call MsecDelay(0.2)

AlcorMPMessage = 0

winHwnd = FindWindow(vbNullString, "AU9540 USBHID AP")

If winHwnd = 0 Then

    Call LoadAP_Click_AU9540
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
    Tester.Print "AP Ready Time="; PassTime

    If PassTime > 5 Then
        TestResult = "Bin3"
        Tester.Label2.Caption = "Bin3"
        Tester.Label2.BackColor = RGB(255, 0, 0)
        Tester.Label9.BackColor = RGB(255, 0, 0)
        Call CloseAU9540AP
        CardResult = DO_WritePort(card, Channel_P1A, &HFF)
        GoTo TestEnd_Label
        'Call MsecDelay(0.2)
        'Exit Sub
    End If

End If

TempRes = WaitDevOn("vid_1059")
Call MsecDelay(0.3)

If TempRes <> 1 Then
    TestResult = "Bin2"
    Tester.Label2.Caption = "Bin2"
    Tester.Label2.BackColor = RGB(255, 0, 0)
    Tester.Label9.Caption = "Find Device Fail"
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
'    AU6256Unknow = AU6256Unknow + 1
'    If AU6256Unknow > 5 Then
'        Shell "cmd /c shutdown -r  -t 0", vbHide
'    End If
'    Call MsecDelay(0.2)
'    Exit Sub
    GoTo TestEnd_Label
End If

Do
    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
    LightCount = LightCount + 1
    Call MsecDelay(0.1)
Loop While (LightOn <> &HFC) And (LightCount < 10)


TempStr = GetDeviceName_NoReply("vid_1059")

Tester.Print "AU9540 begin test........"

If InStr(TempStr, "pid_0018") Then              'HID --- Vid: 1059, Pid: 0018
    Call StartRWTest_Click_AU9540_HID
ElseIf InStr(TempStr, "pid_0017") Then          'CCID -- Vid: 1059, Pid: 0017
    Call StartRWTest_Click_AU9540_CCID
End If

OldTimer = Timer

Do
    If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
        AlcorMPMessage = mMsg.message
        TranslateMessage mMsg
        DispatchMessage mMsg
    End If
         
    PassTime = Timer - OldTimer
       
Loop Until AlcorMPMessage = WM_FT_TEST_UNKNOW _
      Or AlcorMPMessage = WM_FT_TEST_ATR_FAIL _
      Or AlcorMPMessage = WM_FT_TEST_APDU_FAIL _
      Or AlcorMPMessage = WM_FT_TEST_PASS _
      Or PassTime > 5
    
    
Tester.Print "RW work Time="; PassTime
        
'===========================================================
'  RW Time Out Fail
'===========================================================

If PassTime > 5 Then
    TestResult = "Bin3"
    Tester.Label2.Caption = "Bin3"
    Tester.Label9.Caption = "Bin3: Test TimeOut"
    Call CloseAU9540AP
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
    Call MsecDelay(0.2)
    GoTo TestEnd_Label
    'Exit Sub
End If
               
Select Case AlcorMPMessage
            
    Case WM_FT_TEST_UNKNOW
        TestResult = "Bin2"
        Tester.Label2.Caption = "Bin2"
        Tester.Label2.BackColor = RGB(255, 0, 0)
        Tester.Label9.Caption = "Bin2: UnKnow Fail"
'        AU6256Unknow = AU6256Unknow + 1
'        If AU6256Unknow > 5 Then
'            Shell "cmd /c shutdown -r  -t 0", vbHide
'        End If
        Call CloseAU9540AP
        
    Case WM_FT_TEST_ATR_FAIL
        TestResult = "Bin4"
        Tester.Label2.Caption = "Bin4"
        Tester.Label2.BackColor = RGB(255, 0, 0)
        Tester.Label9.Caption = "Bin4: Get ATR Fail"
        AU6256Unknow = 0
        Call CloseAU9540AP
        
    Case WM_FT_TEST_APDU_FAIL
        TestResult = "Bin5"
        Tester.Label2.Caption = "Bin5"
        Tester.Label2.BackColor = RGB(255, 0, 0)
        Tester.Label9.Caption = "Bin5: Get APDU Fail"
        AU6256Unknow = 0
        Call CloseAU9540AP
        
    Case WM_FT_TEST_PASS
        If (LightOn <> &HFC) Then
            TestResult = "Bin3"
            Tester.Label2.Caption = "Bin3"
            Tester.Label2.BackColor = RGB(255, 0, 0)
            Tester.Label9.Caption = "Bin3:LED Fail"
            AU6256Unknow = 0
        Else
            TestResult = "PASS"
            Tester.Label2.Caption = "PASS"
            Tester.Label2.BackColor = RGB(0, 255, 0)
            Tester.Label9.Caption = "PASS"
            AU6256Unknow = 0
        End If
    Case Else
        TestResult = "Bin2"
        Tester.Label2.Caption = "Bin2"
        Tester.Label2.BackColor = RGB(255, 0, 0)
        Tester.Label9.Caption = "Bin2:Undefine Fail"

End Select

TestEnd_Label:




If (Not HV_Done_Flag) Then
    SetSiteStatus (HVDone)
    Call WaitAnotherSiteDone(HVDone, 4#)
    Call PowerSet2(0, "0.0", "0.3", 1, "0.0", "0.3", 1)
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
    WaitDevOFF ("vid_1059")
    Call MsecDelay(0.3)
    HV_Result = TestResult
    HV_Done_Flag = True
    GoTo Routine_Label
Else
    SetSiteStatus (LVDone)
    Call WaitAnotherSiteDone(LVDone, 4#)
    Call PowerSet2(0, "0.0", "0.3", 1, "0.0", "0.3", 1)
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
    WaitDevOFF ("vid_1059")
    Call MsecDelay(0.3)
    LV_Result = TestResult
    
End If


If (HV_Result = "Bin2") And (LV_Result = "Bin2") Then
    TestResult = "Bin2"
    AU6256Unknow = AU6256Unknow + 1
    If AU6256Unknow > 5 Then
        Shell "cmd /c shutdown -r  -t 0", vbHide
    End If
ElseIf (HV_Result <> "PASS") And (LV_Result = "PASS") Then
    AU6256Unknow = 0
    TestResult = "Bin3"
ElseIf (HV_Result = "PASS") And (LV_Result <> "PASS") Then
    AU6256Unknow = 0
    TestResult = "Bin4"
ElseIf (HV_Result <> "PASS") And (LV_Result <> "PASS") Then
    AU6256Unknow = 0
    TestResult = "Bin5"
ElseIf (HV_Result = "PASS") And (LV_Result = "PASS") Then
    AU6256Unknow = 0
    TestResult = "PASS"
Else
    TestResult = "Bin2"
End If


End Sub

Public Sub ReleaseMem(ReleaseMB As Integer)

Dim i As Long
Dim ByteRelease() As Byte

i = CLng(ReleaseMB) * 1024 * 1024
ReDim ByteRelease(i)

Erase ByteRelease

End Sub

Public Sub SmartCard_COM_Mode_TestSub()

Dim OldTimer
Dim PassTime
Dim rt2
Dim mMsg As msg
Dim TempRes As Integer
Dim TempStr As String
Dim LightOn As Long
Dim LightCount As Integer
Dim TargetLEDVal As Long

If PCI7248InitFinish = 0 Then
    PCI7248Exist
End If

If (ChipName = "AU9540CSF20") Then
    TargetLEDVal = 253
ElseIf (ChipName = "AU9540CSF21") Then
    TargetLEDVal = 252
ElseIf (ChipName = "AU9560CSF20") Then
    TargetLEDVal = 252
ElseIf (ChipName = "AU9562CSF20") Then
    TargetLEDVal = 252
End If

If PreChipName <> ChipName Then
    winHwnd = FindWindow(vbNullString, "SmartCard_COM_Mode_Tester")
    If winHwnd <> 0 Then
        Close_SmartCard_COM_Mode
    End If
End If

PreChipName = ChipName

CardResult = DO_WritePort(card, Channel_P1A, &HFC)
Call MsecDelay(0.4)

AlcorMPMessage = 0

winHwnd = FindWindow(vbNullString, "SmartCard_COM_Mode_Tester")

If winHwnd = 0 Then

    Call LoadAP_Click_SmartCard_COM_Mode
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
    Tester.Print "AP Ready Time="; PassTime

    If PassTime > 5 Then
        TestResult = "Bin3"
        Tester.Label2.Caption = "Bin3"
        Tester.Label2.BackColor = RGB(255, 0, 0)
        Tester.Label9.BackColor = RGB(255, 0, 0)
        Call Close_SmartCard_COM_Mode
        CardResult = DO_WritePort(card, Channel_P1A, &HFF)
        Call MsecDelay(0.2)
        Exit Sub
    End If

End If

Tester.Print "Alcor SmartCard Begin Test........"

OldTimer = Timer

Call StartRWTest_Click_SmartCard_COM_Mode

Do
    If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
        AlcorMPMessage = mMsg.message
        TranslateMessage mMsg
        DispatchMessage mMsg
    End If
         
    PassTime = Timer - OldTimer
       
Loop Until AlcorMPMessage = WM_COM_TEST_UNKNOWN _
      Or AlcorMPMessage = WM_COM_TEST_FAIL _
      Or AlcorMPMessage = WM_COM_TEST_PASS _
      Or PassTime > 5
    
    
Tester.Print "RW work Time="; PassTime
        
'===========================================================
'  RW Time Out Fail
'===========================================================

If PassTime > 5 Then
    TestResult = "Bin3"
    Tester.Label2.Caption = "Bin3"
    Tester.Label9.Caption = "Bin3: Test TimeOut"
    Call Close_SmartCard_COM_Mode
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
    Call MsecDelay(0.2)
    Exit Sub
End If

CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
If CardResult <> 0 Then
    MsgBox "Read light On fail"
    End
End If

If LightOn <> TargetLEDVal Then
    TestResult = "Bin5"
    Tester.Label2.Caption = "Bin5"
    Tester.Label9.Caption = "Bin5: LED Fail"
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
    Call MsecDelay(0.2)
    Exit Sub
End If


'MsecDelay (0.2)
               
Select Case AlcorMPMessage
            
    Case WM_COM_TEST_UNKNOWN
        TestResult = "Bin2"
        Tester.Label2.Caption = "Bin2"
        Tester.Label2.BackColor = RGB(255, 0, 0)
        Tester.Label9.Caption = "Bin2: UnKnow Fail"
        
    Case WM_COM_TEST_FAIL
        TestResult = "Bin3"
        Tester.Label2.Caption = "Bin3"
        Tester.Label2.BackColor = RGB(255, 0, 0)
        Tester.Label9.Caption = "Bin3: Get ATR Fail"
        
    Case WM_COM_TEST_PASS
        TestResult = "PASS"
        Tester.Label2.Caption = "PASS"
        Tester.Label2.BackColor = RGB(0, 255, 0)
        Tester.Label9.Caption = "PASS"

    Case Else
        TestResult = "Bin2"
        Tester.Label2.Caption = "Bin2"
        Tester.Label2.BackColor = RGB(255, 0, 0)
        Tester.Label9.Caption = "Bin2:Undefine Fail"

End Select

CardResult = DO_WritePort(card, Channel_P1A, &HFF)
Call MsecDelay(0.2)
    

End Sub
