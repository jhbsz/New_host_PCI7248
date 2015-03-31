Attribute VB_Name = "AU6258Reg"
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Public Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As msg) As Long
Public Declare Function TranslateMessage Lib "user32" (lpMsg As msg) As Long

Public Const WM_USER = &H400
Public Const WM_FT_TEST_READY = WM_USER + &H800
Public Const WM_RADY_TIMER = &H30
Public Const WM_Connect_CK2 = WM_USER + &H100
Public Const WM_Connect_CK2_PASS = WM_USER + &H110
Public Const WM_Connect_CK2_FAIL = WM_USER + &H120

Public Const WM_Start_Test = WM_USER + &H200

Public Const WM_PORT1_FAIL = WM_USER + &H211
Public Const WM_PORT2_FAIL = WM_USER + &H221
Public Const WM_PORT3_FAIL = WM_USER + &H231
Public Const WM_PORT4_FAIL = WM_USER + &H241
Public Const WM_Test_PASS = WM_USER + &H250

Public Const WM_FT_Wifi_FAIL = WM_USER + &H300
Public Const WM_FT_BT_FAIL = WM_USER + &H310
Public Const WM_FT_Cam_FAIL = WM_USER + &H320
Public Const WM_FT_PASS = WM_USER + &H400

Public Const WM_DisConnect_CK2 = WM_USER + &H300
Public Const WM_CLOSE = &H10

Public Const WM_FT_START = WM_USER + &H100
Public Const WM_FT_UNKNOW = WM_USER + &H120
Public Const WM_FT_RW_FAIL = WM_USER + &H200
Public Const WM_FT_CLEAR = WM_USER + &H500


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

Public HiV, LoV As Boolean

Public Type msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Public Sub AU6259_SortingSub()

    If ChipName = "AU6259KFS10" Then
        Call AU6259KFS10_TestSub
    ElseIf ChipName = "AU6259KFS11" Then
        Call AU6259KFS11_TestSub
    ElseIf ChipName = "AU6259KFS20" Then
        Call AU6259KFS20_TestSub
    End If


End Sub


Public Sub AU6258RegTestSub()
    
Dim OldTimer
Dim PassTime
Dim rt2
Dim mMsg As msg
Dim ReTestCount As Integer
    
    rv0 = 0
    rv1 = 0     ' for hiv
    rv2 = 0     ' for lov
    
    If PCI7248InitFinish_Sync = 0 Then
        PCI7248Exist_P1C_Sync
    End If
    
    If AU6258ProgVer <> ChipName Then
        AU6258ProgVer = ChipName
        CloseAU6258FT_Tool
    End If
    
    CardResult = DO_WritePort(card, Channel_P1A, &H0)
    Call MsecDelay(0.2)
    
    Call PowerSet2(0, "5.3", "0.5", 1, "5.3", "0.5", 1)
    Call MsecDelay(0.2)
    
    AlcorMPMessage = 0
    SetSiteStatus (SiteUnknow)

    '===================== Wait AP Ready =====================
    
    winHwnd = FindWindow(vbNullString, "AU6258_RegRead")
    If winHwnd = 0 Then
    
        Call LoadAP_Click_AU6258
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
            Call CloseAU6258FT_Tool
            CardResult = DO_WritePort(card, Channel_P1A, &HFF)
            Call MsecDelay(0.2)
            Exit Sub
        End If
    End If
    
    rv0 = WaitDevOn("3823")
    SetSiteStatus (RunHV)
    Call WaitAnotherSiteDone(RunHV, 2)
    
    If rv0 <> 1 Then
        GoTo AU6258ResultLabel
    End If
    
    Call DisConnect_AU6258_Click
    Call MsecDelay(0.2)
    Call Connect_AU6258_Click
    Call MsecDelay(1#)

    '===================== Find Device =====================
    
    ReTestCount = 0
    
    Do
        rv1 = AU6258_Test()
        ReTestCount = ReTestCount + 1
        If ReTestCount >= 2 Then
            Exit Do
        End If
        
    Loop Until (rv1 = 1)
    
    SetSiteStatus (HVDone)
    If rv1 <> 1 Then
        SetSiteStatus (SiteUnknow)
    End If
    Call WaitAnotherSiteDone(HVDone, 2)
    
    Call PowerSet2(0, "4.7", "0.5", 1, "4.7", "0.5", 1)
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
    Call MsecDelay(0.2)
    CardResult = DO_WritePort(card, Channel_P1A, &H0)
    Call MsecDelay(0.8)
    
    ReTestCount = 0
    Do
        rv2 = AU6258_Test()
        ReTestCount = ReTestCount + 1
        If ReTestCount >= 2 Then
            GoTo AU6258ResultLabel
        End If
        
    Loop Until (rv2 = 1)
    
    SetSiteStatus (RunLV)
    If rv2 <> 1 Then
        SetSiteStatus (SiteUnknow)
    End If
    
AU6258ResultLabel:
    
    Call WaitAnotherSiteDone(RunLV, 2)
    
    Call PowerSet2(0, "0", "0.5", 1, "0", "0.5", 1)
    Call MsecDelay(0.2)

    SetSiteStatus (SiteUnknow)
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
    
    If (rv1 <> 1) Then
        TestResult = "Bin2"
    ElseIf (rv2 <> 1) Then
        TestResult = "Bin3"
    ElseIf (rv1 * rv2 = 1) Then
        TestResult = "PASS"
    End If
    
    If TestResult <> "PASS" Then
        Call CloseAU6258FT_Tool
    End If
    
End Sub

Public Sub AU6257URF20TestSub()
    
Dim OldTimer
Dim PassTime
Dim rt2
Dim mMsg As msg
Dim ReTestCount As Integer
    
    bNeedsReStart = True
    
    rv0 = 0     'Enumeration
    rv1 = 0     'Test result (2:Unknow 3:RW_FAIL)
    
    If PCI7248InitFinish_Sync = 0 Then
        PCI7248Exist_P1C_Sync
    End If
    
'    result = DIO_PortConfig(card, Channel_P1A, INPUT_PORT)
'    Call MsecDelay(0.02)
'    result = DIO_PortConfig(card, Channel_P1A, OUTPUT_PORT)
    
'    If OldChipName <> ChipName Then
'        OldChipName = ChipName
'        CloseAU6257UR_Tool
'    End If
    
    Call AU6257_URTest_ClearBtn
    
    CardResult = DO_WritePort(card, Channel_P1A, &HFE)
    WaitDevOn ("pid_6254")
    Call MsecDelay(0.3)
    CardResult = DO_WritePort(card, Channel_P1A, &H0)
    rv0 = WaitDevOn("vid_067b")
    Call MsecDelay(0.8)
    Call NewLabelMenu(0, "Deice Eum", rv0, 1)
    
    If rv0 <> 1 Then
        GoTo AU6257URTestLabel
    End If
    
    '===================== Wait AP Ready =====================
    
    winHwnd = FindWindow(vbNullString, "Serial_RW")
    If winHwnd = 0 Then
    
        Call LoadAP_Click_AU6257URTest
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
        
        If PassTime >= 5 Then
            Call NewLabelMenu(1, "UART Test", rv1, rv0)
            GoTo AU6257URTestLabel
        End If
        Call MsecDelay(0.2)
    End If

    rv1 = AU6257_StartUARTTest()
    Call NewLabelMenu(1, "UART Test", rv1, rv0)
    
    
AU6257URTestLabel:
    
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
    WaitDevOFF ("vid_067b")
    WaitDevOFF ("pid_6254")
    Call MsecDelay(0.3)
'    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
'    Call MsecDelay(0.3)
    
    If (rv0 <> 1) Then
        TestResult = "Bin2"     'device unknow fail
    ElseIf (rv1 = 2) Then
        TestResult = "Bin3"     'open com-port fail
    ElseIf (rv1 = 3) Then
        TestResult = "Bin4"     'com-port R/W fail
    ElseIf (rv1 <> 1) Then
        TestResult = "Bin5"     'undefine fail
    ElseIf (rv0 * rv1 = 1) Then
        TestResult = "PASS"
    End If
    
    If TestResult <> "PASS" Then
        Call CloseAU6257UR_Tool
    End If
    
End Sub

Public Sub AU6259BFS10TestSub()
    
'This test flow just for Senao HUB Sorting using
Dim OldTimer
Dim PassTime
Dim rt2
Dim mMsg As msg
Dim ReTestCount As Integer
    
    rv0 = 0     'TestResult (1: PASS, 0,2: Fail)
    
    
    If PCI7248InitFinish_Sync = 0 Then
        PCI7248Exist_P1C_Sync
    End If
    
    Call AU6259BF_SenaoSorting_ClearBtn
    
    '===================== Wait AP Ready =====================
    
    winHwnd = FindWindow(vbNullString, "SenaoNetworks_HubSorting")
    If winHwnd = 0 Then
    
        Call LoadAP_Click_AU6259BF_SenaoSorting
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
        
        If PassTime >= 5 Then
            Call NewLabelMenu(1, "Load AP", rv1, 1)
            GoTo AU6259BFSenaoTestLabel
        End If
        Call MsecDelay(0.2)
    End If

    rv1 = AU6259BF_StartSenaoSortingTest()
    Call NewLabelMenu(1, "Format Test", rv1, 1)
    
    
AU6259BFSenaoTestLabel:
    
    
    If (rv1 <> 1) Then
        TestResult = "Bin5"     'fail
    ElseIf (rv1 = 1) Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If

    
End Sub

Public Sub LoadAP_Click_AU6258()
Dim TimePass
Dim rt2
    
    ' find window
    winHwnd = FindWindow(vbNullString, "AU6258_RegRead")
     
    ' run program
    If winHwnd = 0 Then
        If AU6258ProgVer = "AU6258Reg11" Then
            Call ShellExecute(Tester.hwnd, "open", App.Path & "\AU6258\AU6258_RegRead.exe", "400", "", SW_SHOW) 'Connect I2C Speed 400KHz
        Else
            Call ShellExecute(Tester.hwnd, "open", App.Path & "\AU6258\AU6258_RegRead.exe", "", "", SW_SHOW)    'Connect I2C Speed 940KHz
        End If
    End If
    
    SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags
    

End Sub

Public Sub LoadAP_Click_AU6259_SortingTool()
Dim TimePass
Dim rt2
    
    ' find window
    winHwnd = FindWindow(vbNullString, "AU6259_SortingTool")
     
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(Tester.hwnd, "open", App.Path & "\AU6258\AU6259_SortingTool.exe", "", "", SW_SHOW)
    End If
    
    SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags
    

End Sub

Public Sub LoadAP_Click_AU6259_ST2Tool()
Dim TimePass
Dim rt2
    
    ' find window
    winHwnd = FindWindow(vbNullString, "AU6259_SortingTool")
     
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(Tester.hwnd, "open", App.Path & "\AU6258\AU6259_Sorting_BT_HS_FS.exe", "", "", SW_SHOW)
    End If
    
    SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags
    

End Sub

Public Sub LoadAP_Click_AU6257URTest()
Dim TimePass
Dim rt2
    
    ' find window
    winHwnd = FindWindow(vbNullString, "Serial_RW")
     
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(Tester.hwnd, "open", App.Path & "\AU6258\Serial_RW.exe", "", "", SW_SHOW)    'Connect I2C Speed 940KHz
    End If
    
    SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags
    

End Sub

Public Sub LoadAP_Click_AU6259BF_SenaoSorting()
Dim TimePass
Dim rt2
    
    ' find window
    winHwnd = FindWindow(vbNullString, "SenaoNetworks_HubSorting")
     
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(Tester.hwnd, "open", App.Path & "\AU6258\SenaoNetworks_HubSorting.exe", "", "", SW_SHOW)
    End If
    
    SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags
    

End Sub

Public Sub CloseAU6258FT_Tool()
Dim rt2 As Long
Dim EntryTime As Long
Dim PassingTime As Long
Dim mMsg As msg

    winHwnd = FindWindow(vbNullString, "AU6258_RegRead")
    
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
            winHwnd = FindWindow(vbNullString, "AU6258_RegRead")
        Loop While winHwnd <> 0
    End If
    
End Sub

Public Sub CloseAU6257UR_Tool()
Dim rt2 As Long
Dim EntryTime As Long
Dim PassingTime As Long
Dim mMsg As msg

    winHwnd = FindWindow(vbNullString, "Serial_RW")
    
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
            If PassTime > 1 Then
                KillProcess ("Serial_RW")
            Else
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
            End If
    End If
    
End Sub

Public Sub CloseAU6259BF_SenaoSorting_Tool()
Dim rt2 As Long
Dim EntryTime As Long
Dim PassingTime As Long
Dim mMsg As msg

    winHwnd = FindWindow(vbNullString, "SenaoNetworks_HubSorting")
    
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
            winHwnd = FindWindow(vbNullString, "SenaoNetworks_HubSorting")
        Loop While winHwnd <> 0
    End If
    
End Sub

Public Sub CloseAU6259_SortingTool()
Dim rt2 As Long
Dim EntryTime As Long
Dim PassingTime As Long
Dim mMsg As msg

    winHwnd = FindWindow(vbNullString, "AU6259_SortingTool")
    
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
            winHwnd = FindWindow(vbNullString, "AU6258_RegRead")
        Loop While winHwnd <> 0
    End If
    
End Sub

Public Sub DisConnect_AU6258_Click()
Dim rt2
Dim EntryTime As Long
Dim PassingTime As Long
Dim mMsg As msg
    
    winHwnd = FindWindow(vbNullString, "AU6258_RegRead")
    
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
    
        rt2 = PostMessage(winHwnd, WM_DisConnect_CK2, 0&, 0&)
    End If

End Sub

Public Function AU6258_Test() As Byte
Dim rt2
Dim EntryTime As Long
Dim PassingTime As Long
Dim mMsg As msg
    
    AU6258_Test = 0
    
    winHwnd = FindWindow(vbNullString, "AU6258_RegRead")
    
    EntryTime = Timer
    AlcorMPMessage = 0
    
    If winHwnd <> 0 Then
    
        rt2 = PostMessage(winHwnd, WM_Start_Test, 0&, 0&)
        
        Do
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
    
            PassTime = Timer - EntryTime

        Loop Until AlcorMPMessage = WM_Test_PASS _
                Or AlcorMPMessage = WM_PORT1_FAIL _
                Or AlcorMPMessage = WM_PORT2_FAIL _
                Or AlcorMPMessage = WM_PORT3_FAIL _
                Or AlcorMPMessage = WM_PORT4_FAIL _
                Or PassTime > 3
    
    End If
    
    
    If AlcorMPMessage = WM_Test_PASS Then
        AU6258_Test = 1
    End If
    
    Call MsecDelay(0.2)
'
'    If (HiV = True) And (AlcorMPMessage = WM_Test_PASS) Then
'        AU6258_Test = 1
'    ElseIf (HiV = True) And (AlcorMPMessage <> WM_Test_PASS) Then
'        AU6258_Test = 2
'    ElseIf (LoV = True) And (AlcorMPMessage = WM_Test_PASS) Then
'        AU6258_Test = 3
'    ElseIf (LoV = True) And (AlcorMPMessage <> WM_Test_PASS) Then
'        AU6258_Test = 4
'    End If

End Function

Public Function AU6259_StartSortingTest() As Byte
Dim rt2
Dim EntryTime As Long
Dim PassingTime As Long
Dim mMsg As msg
    
    AU6259_StartSortingTest = 0
    
    winHwnd = FindWindow(vbNullString, "AU6259_SortingTool")
    
    EntryTime = Timer
    AlcorMPMessage = 0
    
    If winHwnd <> 0 Then
    
        rt2 = PostMessage(winHwnd, WM_Start_Test, 0&, 0&)
        
        Do
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
    
            PassTime = Timer - EntryTime

        Loop Until AlcorMPMessage = WM_FT_PASS _
                Or AlcorMPMessage = WM_FT_Wifi_FAIL _
                Or AlcorMPMessage = WM_FT_BT_FAIL _
                Or AlcorMPMessage = WM_FT_Cam_FAIL _
                Or PassTime > 15
    
    End If
    
    If AlcorMPMessage = WM_FT_PASS Then
        AU6259_StartSortingTest = 1
    ElseIf AlcorMPMessage = WM_FT_Wifi_FAIL Then    'High Speed Device
        AU6259_StartSortingTest = 2
    ElseIf AlcorMPMessage = WM_FT_BT_FAIL Then
        AU6259_StartSortingTest = 3
    ElseIf AlcorMPMessage = WM_FT_Cam_FAIL Then     'Full Speed Device
        AU6259_StartSortingTest = 4
    Else
        AU6259_StartSortingTest = 0
    End If


End Function

Public Function AU6257_StartUARTTest() As Byte

Dim rt2
Dim EntryTime As Long
Dim PassTime As Long
Dim mMsg As msg
    
    AU6257_StartUARTTest = 0
    
    winHwnd = FindWindow(vbNullString, "Serial_RW")
    
    EntryTime = Timer
    PassTime = 0
    AlcorMPMessage = 0
    
    If winHwnd <> 0 Then
    
        rt2 = PostMessage(winHwnd, WM_FT_START, 0&, 0&)
        
        Do
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
    
            PassTime = Timer - EntryTime

        Loop Until AlcorMPMessage = WM_FT_PASS _
                Or AlcorMPMessage = WM_FT_RW_FAIL _
                Or AlcorMPMessage = WM_FT_UNKNOW _
                Or PassTime > 6
    
    End If
    
    If AlcorMPMessage = WM_FT_PASS Then
        AU6257_StartUARTTest = 1
    ElseIf AlcorMPMessage = WM_FT_UNKNOW Then
        AU6257_StartUARTTest = 2
    ElseIf AlcorMPMessage = WM_FT_RW_FAIL Then
        AU6257_StartUARTTest = 3
    Else
        AU6257_StartUARTTest = 0
    End If

End Function

Public Function AU6259BF_StartSenaoSortingTest() As Byte

Dim rt2
Dim EntryTime As Long
Dim PassTime As Long
Dim mMsg As msg
    
    AU6259BF_StartSenaoSortingTest = 0
    
    winHwnd = FindWindow(vbNullString, "SenaoNetworks_HubSorting")
    
    EntryTime = Timer
    PassTime = 0
    AlcorMPMessage = 0
    
    If winHwnd <> 0 Then
    
        rt2 = PostMessage(winHwnd, WM_FT_START, 0&, 0&)
        
        Do
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
    
            PassTime = Timer - EntryTime

        Loop Until AlcorMPMessage = WM_FT_PASS _
                Or AlcorMPMessage = WM_FT_RW_FAIL _
                Or AlcorMPMessage = WM_FT_UNKNOW _
                Or PassTime > 55
    
    End If
    
    If AlcorMPMessage = WM_FT_PASS Then
        AU6259BF_StartSenaoSortingTest = 1
    ElseIf AlcorMPMessage = WM_FT_RW_FAIL Then
        AU6259BF_StartSenaoSortingTest = 2
    Else
        AU6259BF_StartSenaoSortingTest = 0
    End If

End Function

Public Sub AU6257_URTest_ClearBtn()
Dim rt2
    
    winHwnd = FindWindow(vbNullString, "Serial_RW")
    
    If winHwnd <> 0 Then
        rt2 = PostMessage(winHwnd, WM_FT_CLEAR, 0&, 0&)
    End If
    
End Sub

Public Sub AU6259BF_SenaoSorting_ClearBtn()
Dim rt2
    
    winHwnd = FindWindow(vbNullString, "SenaoNetworks_HubSorting")
    
    If winHwnd <> 0 Then
        rt2 = PostMessage(winHwnd, WM_FT_CLEAR, 0&, 0&)
    End If
    
End Sub

Public Sub Connect_AU6258_Click()
Dim rt2
Dim EntryTime As Long
Dim PassingTime As Long
Dim mMsg As msg
    
    winHwnd = FindWindow(vbNullString, "AU6258_RegRead")
    
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
    
        rt2 = PostMessage(winHwnd, WM_Connect_CK2, 0&, 0&)
        
        Do
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
    
            PassTime = Timer - EntryTime
    
        Loop Until AlcorMPMessage = WM_Connect_CK2_PASS _
                Or AlcorMPMessage = WM_Connect_CK2_FAIL _
                Or PassTime > 1
    End If

End Sub

Public Sub AU6259KFS10_TestSub()
    
Dim OldTimer
Dim PassTime
Dim rt2
Dim mMsg As msg
Dim ReTestCount As Integer
Dim SortingResult As Byte


    If PCI7248InitFinish = 0 Then
        PCI7248Exist
    End If
    
    If AU6258ProgVer <> ChipName Then
        AU6258ProgVer = ChipName
        CloseAU6259_SortingTool
    End If
    
    CardResult = DO_WritePort(card, Channel_P1A, &H0)
    Call MsecDelay(0.2)
    
    Call PowerSet2(0, "3.3", "0.2", 1, "5.0", "0.5", 1)
    Call MsecDelay(0.2)
    
    AlcorMPMessage = 0
    'SetSiteStatus (SiteUnknow)

    '===================== Wait AP Ready =====================
    
    winHwnd = FindWindow(vbNullString, "AU6259_SortingTool")
    If winHwnd = 0 Then
    
        Call LoadAP_Click_AU6259_SortingTool
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
            Call CloseAU6259_SortingTool
            CardResult = DO_WritePort(card, Channel_P1A, &HFF)
            Call MsecDelay(0.2)
            Exit Sub
        End If
    End If
    
    SortingResult = AU6259_StartSortingTest()
    
AU6259ResultLabel:
    
    Call PowerSet2(1, "3.3", "0.2", 1, "0", "0.5", 1)
    Call MsecDelay(0.8)
    Call PowerSet2(1, "0", "0.2", 1, "0", "0.5", 1)
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
    Call MsecDelay(0.6)
    
    If SortingResult = 0 Then
        TestResult = "Bin5"
    ElseIf SortingResult = 2 Then
        TestResult = "Bin2"
    ElseIf SortingResult = 3 Then
        TestResult = "Bin3"
    ElseIf SortingResult = 4 Then
        TestResult = "Bin4"
    ElseIf SortingResult = 1 Then
        TestResult = "PASS"
    Else
        TestResult = "Bin5"
    End If
    
    If TestResult <> "PASS" Then
        Call CloseAU6259_SortingTool
    End If
    
End Sub

Public Sub AU6259KFS11_TestSub()
    
Dim OldTimer
Dim PassTime
Dim rt2
Dim mMsg As msg
Dim ReTestCount As Integer
Dim SortingResult As Byte
'Dim SortingResult_2 As Byte


    'If PCI7248InitFinish = 0 Then
    '    PCI7248Exist
    'End If
    
    If AU6258ProgVer <> ChipName Then
        AU6258ProgVer = ChipName
        CloseAU6259_SortingTool
    End If
    
    SortingResult = 0
    SortingResult_2 = 0
    
    'CardResult = DO_WritePort(card, Channel_P1A, &H0)
    'Call MsecDelay(0.2)
    
    Tester.Print "Begin ST1 test..."
    
    Call PowerSet2(0, "3.3", "0.2", 1, "5.0", "0.5", 1)
    Call MsecDelay(0.2)
    
    AlcorMPMessage = 0
    'SetSiteStatus (SiteUnknow)

    '===================== Wait AP Ready =====================
    
    winHwnd = FindWindow(vbNullString, "AU6259_SortingTool")
    If winHwnd = 0 Then
    
        Call LoadAP_Click_AU6259_SortingTool
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
            Call CloseAU6259_SortingTool
            CardResult = DO_WritePort(card, Channel_P1A, &HFF)
            Call MsecDelay(0.2)
            Exit Sub
        End If
    End If
    
    SortingResult = AU6259_StartSortingTest()
    
'    If SortingResult = 1 Then
'
'        Tester.Print "ST1... PASS"
'
'        Call PowerSet2(1, "3.3", "0.2", 1, "0", "0.5", 1)
'        Call MsecDelay(0.8)
'        Call PowerSet2(1, "0", "0.2", 1, "0", "0.5", 1)
'        Call MsecDelay(0.6)
'
'        Tester.Print "Begin Re-test ST1 test..."
'        Call PowerSet2(0, "3.3", "0.2", 1, "5.0", "0.5", 1)
'        Call MsecDelay(0.2)
'        SortingResult_2 = AU6259_StartSortingTest()
'
'        If SortingResult_2 = 1 Then
'            Tester.Print "ST1 Re-Test... PASS"
'        Else
'            Tester.Print "ST1 Re-Test... Fail"
'        End If
'
'    Else
'        Tester.Print "ST1... Fail"
'    End If
    
    
AU6259ResultLabel:
    
    Call PowerSet2(1, "3.3", "0.2", 1, "0", "0.5", 1)
    Call MsecDelay(0.8)
    Call PowerSet2(1, "0", "0.2", 1, "0", "0.5", 1)
    'CardResult = DO_WritePort(card, Channel_P1A, &HFF)
    Call MsecDelay(0.6)
    
    If SortingResult = 2 Then
        TestResult = "Bin2"
    ElseIf SortingResult = 3 Then
        TestResult = "Bin3"
    ElseIf SortingResult = 4 Then
        TestResult = "Bin4"
    ElseIf SortingResult = 1 Then
        TestResult = "PASS"
    End If
    
'    If TestResult <> "PASS" Then
'        Call CloseAU6259_SortingTool
'    End If
    
End Sub

Public Sub AU6259KFS20_TestSub()
    
Dim OldTimer
Dim PassTime
Dim rt2
Dim mMsg As msg
Dim ReTestCount As Integer
Dim SortingResult As Byte
'Dim SortingResult_2 As Byte


    If AU6258ProgVer <> ChipName Then
        AU6258ProgVer = ChipName
        CloseAU6259_SortingTool
    End If
    
    SortingResult = 0
    SortingResult_2 = 0
    
    'CardResult = DO_WritePort(card, Channel_P1A, &H0)
    'Call MsecDelay(0.2)
    
    Tester.Print "Begin ST2 test..."
    
    Call PowerSet2(0, "3.3", "0.2", 1, "5.0", "0.5", 1)
    Call MsecDelay(3.2)
    
    AlcorMPMessage = 0
    'SetSiteStatus (SiteUnknow)

    '===================== Wait AP Ready =====================
    
    winHwnd = FindWindow(vbNullString, "AU6259_SortingTool")
    If winHwnd = 0 Then
    
        Call LoadAP_Click_AU6259_ST2Tool
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
            Call CloseAU6259_SortingTool
            CardResult = DO_WritePort(card, Channel_P1A, &HFF)
            Call MsecDelay(0.2)
            Exit Sub
        End If
    End If
    
    SortingResult = AU6259_StartSortingTest()
    
    
AU6259ResultLabel:
    
    Call PowerSet2(1, "3.3", "0.2", 1, "0.0", "0.5", 1)
    Call MsecDelay(0.8)
    Call PowerSet2(1, "0", "0.2", 1, "0.0", "0.5", 1)
    'CardResult = DO_WritePort(card, Channel_P1A, &HFF)
    Call MsecDelay(0.6)
    
    If SortingResult = 2 Then       'Elecom Fail
        TestResult = "Bin2"
    ElseIf SortingResult = 3 Then   'BT Fail
        TestResult = "Bin3"
    ElseIf SortingResult = 4 Then   'AU9368 Reader Fail
        TestResult = "Bin4"
    ElseIf SortingResult = 1 Then
        TestResult = "PASS"
    End If
    
'    If TestResult <> "PASS" Then
'        Call CloseAU6259_SortingTool
'    End If
    
End Sub
