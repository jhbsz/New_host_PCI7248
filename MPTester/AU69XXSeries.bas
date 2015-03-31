Attribute VB_Name = "AU69XXSeries"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" _
   Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
   ByVal lpKeyName As Any, ByVal lpDefault As String, _
   ByVal lpReturnedString As String, ByVal nSize As Long, _
   ByVal lpFileName As String) As Long

' INI file
Public Declare Function WritePrivateProfileString Lib "kernel32" _
   Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
   ByVal lpKeyName As Any, ByVal lpString As Any, _
   ByVal lpFileName As String) As Long

'========= Define AnOther Site State =========
' PortC Configuration:
' 1. GPIB Exist: Hi-Byte => Output
'                Li-Byte => Input
'
' 2. No GPIB   : Hi-Byte => Input
'                Li-Byte => Output

Public Const SiteReady = &H0
Public Const RunMP = &H1
Public Const MPDone = &H2
Public Const RunHV = &H3
Public Const HVDone = &H4
Public Const RunLV = &H5
Public Const LVDone = &H6
Public Const SiteUnknow = &HF
Public Const RunU2 = &H7
Public Const U2Done = &H8
Public Const RunU3 = &H9
Public Const U3Done = &HA

Public CheckLED_Flag As Boolean

'========== AU6988 message
Public Const WM_USER = &H400
Public Const WM_FT_MP_START = WM_USER + &H20
Public Const WM_FT_MP_STOP = WM_USER + &H21
Public Const WM_FT_MP_REFRESH = WM_USER + &H22

Public Const WM_FT_U3_MP_LLF = WM_USER + &H26
Public Const WM_FT_U3_MP_RESET = WM_USER + &H27

Public Const WM_FT_MP_PASS = WM_USER + &H40
Public Const WM_FT_MP_UNKNOW_FAIL = WM_USER + &H60
Public Const WM_FT_MP_FAIL = WM_USER + &H80

Public Const WM_FT_RW_START = WM_USER + &H100
Public Const WM_FT_RW_UNKNOW_FAIL = WM_USER + &H120
Public Const WM_FT_LOCK_VOLUME_FAIL = WM_USER + &H130
Public Const WM_FT_UNLOCK_VOLUME_FAIL = WM_USER + &H140
Public Const WM_FT_TURNOFF_TUR_FAIL = WM_USER + &H150
Public Const WM_FT_RW_SPEED_FAIL = WM_USER + &H200
Public Const WM_FT_RW_RW_FAIL = WM_USER + &H201
Public Const WM_FT_RW_ROM_FAIL = WM_USER + &H210
Public Const WM_FT_RW_RAM_FAIL = WM_USER + &H220
Public Const WM_FT_CHECK_CERBGPO_FAIL = WM_USER + &H230

Public Const WM_FT_CHECK_HW_CODE_FAIL = WM_USER + &H240

Public Const WM_FT_PHYREAD_FAIL = WM_USER + &H250
Public Const WM_FT_ECC_FAIL = WM_USER + &H260
Public Const WM_FT_NOFREEBLOCK_FAIL = WM_USER + &H270
Public Const WM_FT_LODECODE_FAIL = WM_USER + &H280
Public Const WM_LC_FAIL = WM_USER + &H290
Public Const WM_RC_FAIL = WM_USER + &H300

'Public Const WM_FT_RELOADCODE_FAIL = WM_USER + &H290
'Public Const WM_FT_TESTUNITREADY_FAIL = WM_USER + &H300

Public Const WM_FT_CHECK_WRITE_PROTECT_FAIL = WM_USER + &H310
Public Const WM_FT_NO_CARD_FAIL = WM_USER + &H320
Public Const WM_FT_MOVE_DATA_FAIL = WM_USER + &H330
Public Const WM_FT_RW_RW_PASS = WM_USER + &H400
Public Const WM_FT_FLASH_NUM_FAIL = WM_USER + &H410

'Public Const WM_FT_READ_MANY_TIME_FAIL = WM_USER + &H420

Public Const WM_FT_TEST_DQS_FAIL = WM_USER + &O430
Public Const WM_FT_RW_READY = WM_USER + &H800

'Public Const WM_READY_TIMER = WM_USER + &H900

Public Const WM_DEV_GET_HANDLE_FAIL = WM_USER + &H910
Public Const WM_DEV_GET_DIS_TUR_HANDLE_FAIL = WM_USER + &H920
Public Const WM_SEND_DEV_READY = WM_USER + &H930


' === AU6996 reader message
Public Const WM_FT_PARAM_FAIL = WM_USER + &H130
Public Const WM_FT_READER_FAIL = WM_USER + &H140
Public Const WM_FT_BUSWIDTH_FAIL = WM_USER + &H150
Public Const WM_FT_BUSCLK_FAIL = WM_USER + &H160

' === AU87100 U3 ===
Public Const WM_FT_TEST_U2 = WM_USER + &H580
Public Const WM_FT_TEST_U3 = WM_USER + &H590
Public U2_Pass As Boolean
Public U3_Pass As Boolean
Public U3_Test As Boolean

' === AU6928 new ===
Public Const WM_FT_MP_RECOGNIZE = WM_USER + &H25
Public Const WM_FT_MP_UNRECOGNIZE = WM_USER + &H26

Public AlcorMPHandler As Long
Public AlcorMPMessage As Long

Public Const AU6997MPCaption1 = "698x UFD MP"                       'CurdevicePar.MP_LoadTitle
Public Const AU6997MPCaption = "698x UFD MP, Cycle Time : 50 ns"    'CurdevicePar.MP_Work_Title
Public Const AU6996MPCaption1 = "698x UFD MP"
Public Const AU6996MPCaption = "698x UFD MP, Cycle Time : 33 ns"
Public Const AU6992MPCaption1 = "698x UFD MP"
Public Const AU6992MPCaption = "698x UFD MP, Cycle Time : 33 ns"

Public Const AU6988MPCaption1 = "698x UFD MP"
Public Const AU6988MPCaption = "698x UFD MP, Cycle Time : 33 ns"

Private Type DeviceParFormat
    Dual_Flag As Boolean
    EQC_Flag As Boolean
    FullName As String
    ShortName As String
    SetStdV1 As String
    SetStdV2 As String
    SetHLVStd1 As String
    SetHLVStd2 As String
    SetHV1 As String
    SetHV2 As String
    SetLV1 As String
    SetLV2 As String
    SetStdI1 As String
    SetStdI2 As String
    SetHI1 As String
    SetHI2 As String
    SetLI1 As String
    SetLI2 As String
    MP_ToolFileName As String
    MP_LoadTitle As String
    MP_WorkTitle As String
    DeviceFolder As String
'    UpdateModuleName As String
    FT_ToolFileName As String
    FT_ToolTitle As String
    Exec_Par1 As String
    Exec_Par2 As String
End Type

Public CurDevicePar As DeviceParFormat

Public FailCloseAP As Boolean

Dim HV_Result As String
Dim LV_Result As String

Public OldVer_Flag As Boolean
Public ST4FirstMP As Boolean

Public bJitterSorting As Boolean

Public Function GetAnOtherStatus() As Byte

Dim StatusInCH As Integer
Dim StatusVal As Long
    
    GetAnOtherStatus = &HF
    
    If GPIBCard_Exist Then
        StatusInCH = Channel_P1CL
    Else
        StatusInCH = Channel_P1CH
    End If
    
    'DoEvents
    cardresult = DO_ReadPort(card, StatusInCH, StatusVal)
    GetAnOtherStatus = StatusVal
    'debug.print GetAnOtherStatus
    
End Function

Public Sub SetSiteStatus(SetVal As Long)

Dim StatusOutCH As Integer

    If GPIBCard_Exist Then
        StatusOutCH = Channel_P1CH
    Else
        StatusOutCH = Channel_P1CL
    End If
    
    'DoEvents
    cardresult = DO_WritePort(card, StatusOutCH, SetVal)
    
End Sub

Public Sub WaitAnotherSiteDone(FlowItem As Long)

Dim ReadVal As Long
Dim PassTime As Single
Dim OldTimer As Single

    OldTimer = Timer

    Do
        ReadVal = GetAnOtherStatus
        If (ReadVal = SiteUnknow) Or (ReadVal = FlowItem) Then
            Exit Sub
        End If
        Call MsecDelay(0.06)
        PassTime = Timer - OldTimer
        
        If (FlowItem <> MPDone) And (PassTime > 5) Then
            Exit Sub
        End If
    
    Loop Until (ReadVal = FlowItem) Or (PassTime > 30)
    
End Sub

Public Sub LoadMP_AU69XX()

Dim TimePass
Dim rt2

    If (CurDevicePar.DeviceFolder = "") Or (CurDevicePar.MP_ToolFileName = "") Or (CurDevicePar.MP_LoadTitle = "") Then  ' (CurDevicePar.UpdateModuleName = "") Or _

        'MsgBox ("DeviceFolder/UpdateModuleName/MP_ToolFileName Unknow")
        MsgBox ("DeviceFolder/MP_ToolFileName Unknow")
        Exit Sub
    End If
 
    winHwnd = FindWindow(vbNullString, CurDevicePar.MP_WorkTitle)
    
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & CurDevicePar.DeviceFolder & "\" & CurDevicePar.MP_ToolFileName, "", "", SW_SHOW)
    End If
    
    SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags

End Sub

Public Sub AU87100_LLF()

Dim rt2

    If (CurDevicePar.MP_WorkTitle = "") Then
        MsgBox ("MP_WorkTitle UNknow")
        Exit Sub
    End If
    
    winHwnd = FindWindow(vbNullString, CurDevicePar.MP_WorkTitle)
    rt2 = PostMessage(winHwnd, WM_FT_U3_MP_LLF, 0&, 0&)
    
End Sub

Public Sub StartMP_AU69XX()

Dim rt2

    If (CurDevicePar.MP_WorkTitle = "") Then
        MsgBox ("MP_WorkTitle UNknow")
        Exit Sub
    End If
    
    winHwnd = FindWindow(vbNullString, CurDevicePar.MP_WorkTitle)
    rt2 = PostMessage(winHwnd, WM_FT_MP_START, 0&, 0&)
    
End Sub

Public Sub RefreshMP_AU69XX()

Dim rt2
    
    If (CurDevicePar.MP_WorkTitle = "") Then
        MsgBox ("MP_WorkTitle UNknow")
        Exit Sub
    End If
    
    winHwnd = FindWindow(vbNullString, CurDevicePar.MP_WorkTitle)
    rt2 = PostMessage(winHwnd, WM_FT_MP_REFRESH, 0&, 0&)
    
End Sub

Public Sub CloseMP_AU69XX()

Dim rt2

    If (CurDevicePar.MP_LoadTitle = "") Or (CurDevicePar.MP_WorkTitle = "") Then
        MsgBox ("MP_WorkTitle/MP_LoadTitle UNknow")
        Exit Sub
    End If

    '(1)
    winHwnd = FindWindow(vbNullString, CurDevicePar.MP_LoadTitle)
    
    If winHwnd <> 0 Then
        Do
            rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
            Call MsecDelay(0.5)
            winHwnd = FindWindow(vbNullString, CurDevicePar.MP_LoadTitle)
        Loop While winHwnd <> 0
    End If
    
    '(2)
    winHwnd = FindWindow(vbNullString, CurDevicePar.MP_WorkTitle)
    
    If winHwnd <> 0 Then
        Do
            rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
            Call MsecDelay(0.5)
            winHwnd = FindWindow(vbNullString, CurDevicePar.MP_WorkTitle)
        Loop While winHwnd <> 0
    End If
    
    '(3)
    winHwnd = FindWindow(vbNullString, "重新啟動量產程式")

    If winHwnd <> 0 Then
        Do
            rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
            Call MsecDelay(0.5)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
        Loop While winHwnd <> 0
    End If
    
End Sub

Public Sub UnloadDriver()

Dim pid As Long          ' unload driver
Dim hProcess As Long
Dim ExitEvent As Long

    pid = Shell(App.Path & "\LoadDrv.exe uninstall_058F6387")
    hProcess = OpenProcess(SYNCHRONIZE + PROCESS_QUERY_INFORMATION + PROCESS_TERMINATE, 0, pid)
    ExitEvent = WaitForSingleObject(hProcess, INFINITE)
    Call CloseHandle(hProcess)
    KillProcess ("LoadDrv.exe")

End Sub

Public Sub LoadFTtool_AU69XX()

    'find window
    winHwnd = FindWindow(vbNullString, CurDevicePar.FT_ToolTitle)
 
    'run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & CurDevicePar.DeviceFolder & "\" & CurDevicePar.FT_ToolFileName, CurDevicePar.Exec_Par1, "", SW_SHOW)
    End If
 
End Sub

Public Sub CloseFTtool_AU69XX()

    winHwnd = FindWindow(vbNullString, CurDevicePar.FT_ToolTitle)
    
    If winHwnd <> 0 Then
        Do
            winHwnd = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
            Call MsecDelay(0.2)
            winHwnd = FindWindow(vbNullString, CurDevicePar.FT_ToolTitle)
        Loop While winHwnd <> 0
    End If
    
End Sub

Public Sub StartFTtest_WaitDevReady_AU69XX(TimeOut As Single, Optional MPFlag As Boolean)
Dim EntryTime As Long
Dim PassTime As Long
Dim mMsg As MSG
Dim rt2

    EntryTime = Timer

    Do                                                   'Wait AP send Dev_Ready message
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
        
        PassTime = Timer - EntryTime
        
    Loop Until AlcorMPMessage = WM_SEND_DEV_READY _
            Or AlcorMPMessage = WM_FT_RW_UNKNOW_FAIL _
            Or PassTime > TimeOut _
            Or AlcorMPMessage = WM_CLOSE _
            Or AlcorMPMessage = WM_DESTROY

    'If PassTime <= TimeOut Then
    If AlcorMPMessage = WM_SEND_DEV_READY Then
        Call MsecDelay(0.35)
    End If
    
    If MPFlag Then
        Call MsecDelay(11#)
    End If

    winHwnd = FindWindow(vbNullString, CurDevicePar.FT_ToolTitle)

    rt2 = PostMessage(winHwnd, WM_FT_RW_START, 0&, 0&)

    
End Sub

Public Sub StartFTtest_AU69XX()
Dim rt2

    winHwnd = FindWindow(vbNullString, CurDevicePar.FT_ToolTitle)
    rt2 = PostMessage(winHwnd, WM_FT_RW_START, 0&, 0&)

End Sub

Public Sub AU6928XXXHLS50TestSub()

Dim PassTime
Dim OldTimer
Dim mMsg As MSG

    If PCI7248InitFinish = 0 Then
       PCI7248Exist
    End If

    cardresult = DO_WritePort(card, Channel_P1A, &HFF)
    Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
    WaitDevOFF ("vid_058f")
    MsecDelay (0.3)
    cardresult = DO_WritePort(card, Channel_P1A, &HFB)
    Call PowerSet2(1, "5.0", "0.5", 1, "5.0", "0.5", 1)
    WaitDevOn ("vid_058f")

    MsecDelay (0.1)
    
    winHwnd = FindWindow(vbNullString, CurDevicePar.FT_ToolTitle)

    If winHwnd <> 0 Then
        Call CloseFTtool_AU69XX
    End If
    
    Call LoadMP_AU69XX

    AlcorMPMessage = 0
    OldTimer = Timer

    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If

        PassTime = Timer - OldTimer

    Loop Until AlcorMPMessage = WM_FT_MP_RECOGNIZE Or AlcorMPMessage = WM_FT_MP_UNRECOGNIZE Or PassTime > 5
    
    Call Check_STD_RW_Result(PassTime)
    
End Sub

Public Sub AU69XX_StdTestSub()

Dim OldTimer
Dim PassTime
Dim mMsg As MSG
Dim ICName As String, MODELName As String
Dim tmpName As String
Dim TimerCounter As Integer
Dim TmpString As String
Dim MP_Retry As Byte
Dim RWDone As Boolean
Dim refreshMSG


    'add unload driver function
    If PCI7248InitFinish = 0 Then
       PCI7248Exist
    End If

    PLFlag = False  ' for AU6998-PL SOCKET
    
    KLFlag = False  ' for AU6927-KL
    
    If InStr(ChipName, "PL") Then
        PLFlag = True
    End If
    
    If InStr(ChipName, "KL") Then
        KLFlag = True
    End If

    MP_Retry = 0
    
    If Len(ChipName) = 15 Then
        tmpName = Left(ChipName, 10)
    Else
        tmpName = Left(ChipName, 9)
    End If
    
    MPTester.TestResultLab = ""
    NewChipFlag = 0
    ChDir (App.Path & CurDevicePar.DeviceFolder)
        
    If OldChipName <> ChipName Then
        Call Fail_Location_Initial(tmpName)
        
        Call CloseFTtool_AU69XX
        
        If Left(ChipName, 6) = "AU6921" Or Left(ChipName, 6) = "AU692A" Or Left(ChipName, 6) = "AU6928" Or Left(ChipName, 6) = "AU692H" Or Left(ChipName, 6) = "AU692B" Then
            ' for MSL
            If (InStr(ChipName, "SLF")) Then
                FileCopy App.Path & CurDevicePar.DeviceFolder & "\62026UFDTest.exe", App.Path & CurDevicePar.DeviceFolder & "\UFDTest.exe"
            Else
                FileCopy App.Path & CurDevicePar.DeviceFolder & "\60026UFDTest.exe", App.Path & CurDevicePar.DeviceFolder & "\UFDTest.exe"
            End If
        End If
        
        If Left(ChipName, 6) = "AU6930" Then
            ' for MSL
            If (InStr(ChipName, "SLF")) Then
                FileCopy App.Path & CurDevicePar.DeviceFolder & "\62028UFDTest.exe", App.Path & CurDevicePar.DeviceFolder & "\UFDTest.exe"
            Else
                FileCopy App.Path & CurDevicePar.DeviceFolder & "\60028UFDTest.exe", App.Path & CurDevicePar.DeviceFolder & "\UFDTest.exe"
            End If
        End If
        
    End If

    OldChipName = ChipName
    SetSiteStatus (SiteReady)
    
    '==============================================================
    ' when begin RW Test, must clear MP program
    '===============================================================
    Call CloseMP_AU69XX
    MPTester.Print "ContFail="; ContFail
    MPTester.Print "MPContFail="; MPContFail
 
    '====================================
    '  Fix Card
    '====================================
    If InStr(ChipName, "87100") Then
        ContFail = 1
        U3MPFlag = U3MPFlag + 1
        Debug.Print U3MPFlag
    End If
    
    
    If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag) Or (U3MPFlag = 50) Then
        If MPTester.NoMP.Value = 1 Then
            If (NewChipFlag = 0) And (MPTester.Check1.Value = 0) Then  ' force condition
                'Call STD_RW_TEST(PassTime)
                RWDone = True
                'GoTo RW_Test_Label
            End If
        End If
        
        If RWDone = False Then
            If MPTester.ResetMPFailCounter.Value = 1 Then
                ContFail = 0
            End If
            
            '===============================================================
            ' when begin MP, must close RW program
            '===============================================================
            U3MPFlag = 0
            MPFlag = 1
            U2_Pass = False
            Call CloseFTtool_AU69XX
            
            'power off
            If InStr(ChipName, "87100") Then
                cardresult = DO_WritePort(card, Channel_P1A, &H1)
            Else
                cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'close power
            End If
            Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)   ' close power to disable chip
            Call MsecDelay(0.5)  ' power for load MPDriver
            MPTester.Print "wait for MP Ready"
            Call LoadMP_AU69XX
        
            OldTimer = Timer
            AlcorMPMessage = 0
            Do
                If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                    AlcorMPMessage = mMsg.message
                    TranslateMessage mMsg
                    DispatchMessage mMsg
                End If
                PassTime = Timer - OldTimer
            
            Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 30 _
                    Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
        
            MPTester.Print "Ready Time="; PassTime
            
            '====================================================
            '  handle MP load time out, the FAIL will be Bin3
            '====================================================
            If PassTime > 30 Then
                Call CloseMP_AU69XX
                MPTester.TestResultLab = "Bin3:MP Ready Fail"
                TestResult = "Bin3"
                MPTester.Print "MP Ready Fail"
                SetSiteStatus (SiteUnknow)
                Call TestEnd
                Exit Sub
            End If
                
            '====================================================
            '  MP begin
            '====================================================
            
            If AlcorMPMessage = WM_FT_MP_START Then
                
                SetSiteStatus (RunMP)
                WaitAnotherSiteDone (RunMP)
                
                If Dir("D:\LABPC.PC") = "LABPC.PC" Then
                    Call PowerSet2(1, "5.0", "0.5", 1, "5.0", "0.5", 1)
                Else
                    Call PowerSet2(1, CurDevicePar.SetStdV1, "0.5", 1, CurDevicePar.SetStdV2, "0.5", 1)
                End If
                
                If Dual_Flag Then
                    cardresult = DO_WritePort(card, Channel_P1A, &HFD)
                Else
                    If InStr(ChipName, "87100") Then
                        If U2_Pass = False Then
                            cardresult = DO_WritePort(card, Channel_P1A, &HF6)
                        Else
                            cardresult = DO_WritePort(card, Channel_P1A, &HF5)
                            MsecDelay (0.05)
                            cardresult = DO_WritePort(card, Channel_P1A, &HF4)
                        End If
                    Else
                        cardresult = DO_WritePort(card, Channel_P1A, &HFB)
                    End If
                End If
                
                Call MsecDelay(0.5)
                TimerCounter = 0
                
                Do
                    DoEvents
                    Call MsecDelay(0.1)
                    TimerCounter = TimerCounter + 1

                    TmpString = GetDeviceName_NoReply("058f")
                    If TmpString = "" Then
                        TmpString = GetDeviceName_NoReply("8564")
                    End If
                    

                Loop While (TmpString = "") And (TimerCounter < 100)
                     
                winHwnd = FindWindow(vbNullString, CurDevicePar.MP_WorkTitle)
                Call MsecDelay(0.1)
                refreshMSG = PostMessage(winHwnd, WM_FT_MP_REFRESH, 0&, 0&)
                Call MsecDelay(2#)
                
                If MPTester.LLF.Value = 1 Then
                    Call AU87100_LLF
                    OldTimer = Timer
                    AlcorMPMessage = 0
                    ReMP_Flag = 0
                    Call MsecDelay(2#)
                    
                    Do
                        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                            AlcorMPMessage = mMsg.message
                            TranslateMessage mMsg
                            DispatchMessage mMsg
                            
                            If (AlcorMPMessage = WM_FT_MP_FAIL) And (MP_Retry < 3) Then
                                AlcorMPMessage = 1
                                If InStr(ChipName, "87100") Then
                                    cardresult = DO_WritePort(card, Channel_P1A, &H1)   'close power
                                Else
                                    cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'close power
                                End If
                                
                                Call MsecDelay(0.3)
                                If Dual_Flag Then
                                    cardresult = DO_WritePort(card, Channel_P1A, &HFD)
                                Else
                                    If InStr(ChipName, "87100") Then
                                        If U2_Pass = False Then
                                            cardresult = DO_WritePort(card, Channel_P1A, &HF6)
                                        Else
                                            cardresult = DO_WritePort(card, Channel_P1A, &HF5)
                                            MsecDelay (0.05)
                                            cardresult = DO_WritePort(card, Channel_P1A, &HF4)
                                        End If
                                    Else
                                        cardresult = DO_WritePort(card, Channel_P1A, &HFB)
                                    End If
                                End If
                                Call MsecDelay(2.2)
                                Call RefreshMP_AU69XX
                                Call MsecDelay(0.5)
                                Call AU87100_LLF
                                MP_Retry = MP_Retry + 1
                            End If
                        End If
                        
                        PassTime = Timer - OldTimer
                    Loop Until AlcorMPMessage = WM_FT_MP_PASS _
                            Or AlcorMPMessage = WM_FT_MP_FAIL _
                            Or AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL _
                            Or PassTime > 65 _
                            Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                End If
                
                    
                If TmpString = "" Then   ' can not find device after 15 s
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail when enter MP"
                    SetSiteStatus (SiteUnknow)
                    Call TestEnd
                    Exit Sub
                End If
                     
                Call MsecDelay(2.5)
                
                winHwnd = FindWindow(vbNullString, CurDevicePar.MP_WorkTitle)
                Call MsecDelay(0.1)
                refreshMSG = PostMessage(winHwnd, WM_FT_MP_REFRESH, 0&, 0&)
                Call MsecDelay(2#)
                
                MPTester.Print " MP Begin....."
                     
                Call StartMP_AU69XX
                OldTimer = Timer
                AlcorMPMessage = 0
                ReMP_Flag = 0
                Call MsecDelay(2#)
                
                Do
                    If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                        AlcorMPMessage = mMsg.message
                        TranslateMessage mMsg
                        DispatchMessage mMsg
                        
                        If (AlcorMPMessage = WM_FT_MP_FAIL) And (MP_Retry < 3) Then
                            AlcorMPMessage = 1
                            If InStr(ChipName, "87100") Then
                                cardresult = DO_WritePort(card, Channel_P1A, &H1)   'close power
                            Else
                                cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'close power
                            End If
                            
                            Call MsecDelay(0.3)
                            If Dual_Flag Then
                                cardresult = DO_WritePort(card, Channel_P1A, &HFD)
                            Else
                                If InStr(ChipName, "87100") Then
                                    If U2_Pass = False Then
                                        cardresult = DO_WritePort(card, Channel_P1A, &HF6)
                                    Else
                                        cardresult = DO_WritePort(card, Channel_P1A, &HF5)
                                        MsecDelay (0.05)
                                        cardresult = DO_WritePort(card, Channel_P1A, &HF4)
                                    End If
                                Else
                                    cardresult = DO_WritePort(card, Channel_P1A, &HFB)
                                End If
                            End If
                            Call MsecDelay(2.2)
                            Call RefreshMP_AU69XX
                            Call MsecDelay(0.5)
                            Call StartMP_AU69XX
                            MP_Retry = MP_Retry + 1
                        End If
                    End If
                    
                    PassTime = Timer - OldTimer
                Loop Until AlcorMPMessage = WM_FT_MP_PASS _
                        Or AlcorMPMessage = WM_FT_MP_FAIL _
                        Or AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL _
                        Or PassTime > 65 _
                        Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                        
                MPTester.Print "MP work time="; PassTime
                MPTester.MPText.Text = Hex(AlcorMPMessage)
                MsecDelay (0.5)
                         
                '=====================
                '    Close MP program
                '=====================
                Call CloseMP_AU69XX
                KillProcess ("AlcorMP.exe")
                Call UnloadDriver
                
                '================================
                '  Handle MP work time out error
                '================================
                'Call Check_MP_Result(PassTime)
                
                If PassTime > 65 Then
                    Call CloseMP_AU69XX
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    SetSiteStatus (SiteUnknow)
                    Call TestEnd
                    Exit Sub
                End If
                   
                'MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL Then
                    Call CloseMP_AU69XX
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    SetSiteStatus (SiteUnknow)
                    Call TestEnd
                    Exit Sub
                End If
            
                'unknow fail
                If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                    Call CloseMP_AU69XX
                    MPContFail = MPContFail + 1
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                    MPTester.Print "MP UNKNOW Fail"
                    SetSiteStatus (SiteUnknow)
                    Call TestEnd
                    Exit Sub
                End If
                   
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_PASS Then
                    MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPTester.Print "MP PASS"
                    SetSiteStatus (MPDone)
                End If
            End If
        End If
    End If
       
'        If bJitterSorting Then
'            winHwnd = FindWindow(vbNullString, "BurnInTest V4.0 Pro - [Live Results]")
'
'            If winHwnd = 0 Then
'                ' run program
'                If bJitterSorting Then
'                    Shell App.Path & "\BurnInTestV40\bit.exe -r", vbNormalFocus
'                    Call MsecDelay(1#)
'                End If
'            End If
'        End If
       
    If PassTime < 65 And (AlcorMPMessage <> WM_FT_MP_FAIL Or AlcorMPMessage <> WM_FT_MP_UNKNOW_FAIL) Then
        If STD_RW_TEST <= 10 Then                 ' TimeOut = 8
            Call Check_STD_RW_Result(PassTime)
            SetSiteStatus (HVDone)
            WaitAnotherSiteDone (HVDone)
        End If
        
        Call TestEnd
    End If
    
End Sub

Public Sub Fail_Location_Initial(tmpName As String)

'    If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
'        FileCopy App.Path & CurDevicePar.DeviceFolder & "\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
'        Call MsecDelay(5)
'    End If

    If (OldVer_Flag) And (InStr(tmpName, "6927") = 0) Then
        FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\ROM.Hex", App.Path & CurDevicePar.DeviceFolder & "\ROM.Hex"
        FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\RAM.Bin", App.Path & CurDevicePar.DeviceFolder & "\RAM.Bin"
        
        If PLFlag = True Then
            FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\AlcorMP_PL.ini", App.Path & CurDevicePar.DeviceFolder & "\AlcorMP.ini"
        ElseIf KLFlag = True Then
            FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\AlcorMP_KL.ini", App.Path & CurDevicePar.DeviceFolder & "\AlcorMP.ini"
        Else
            FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\AlcorMP.ini", App.Path & CurDevicePar.DeviceFolder & "\AlcorMP.ini"
        End If
        
        FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\PE.bin", App.Path & CurDevicePar.DeviceFolder & "\PE.bin"
        FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\FT.ini", App.Path & CurDevicePar.DeviceFolder & "\FT.ini"
    Else
    
        If (InStr(tmpName, "AU692H")) Then
            tmpName = Replace(tmpName, "2H", "28")
        End If
    
        FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\ROM.Hex", App.Path & CurDevicePar.DeviceFolder & "\ROM.Hex"
        FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\RAM.Bin", App.Path & CurDevicePar.DeviceFolder & "\RAM.Bin"
        
        'If (InStr(tmpName, "PL")) Then
        '    FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\AlcorMP_PL.ini", App.Path & CurDevicePar.DeviceFolder & "\AlcorMP.ini"
        'Else
            FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\AlcorMP.ini", App.Path & CurDevicePar.DeviceFolder & "\AlcorMP.ini"
        'End If
        
        FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\PE.bin", App.Path & CurDevicePar.DeviceFolder & "\PE.bin"
        
        If PLFlag = True Then
            FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\PL_FT.ini", App.Path & CurDevicePar.DeviceFolder & "\FT.ini"
        ElseIf KLFlag = True Then
            FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\KL_FT.ini", App.Path & CurDevicePar.DeviceFolder & "\FT.ini"
        Else
            FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\FT.ini", App.Path & CurDevicePar.DeviceFolder & "\FT.ini"
        End If
        
    End If
    
    NewChipFlag = 1 ' force MP
End Sub

Public Sub Check_MP_Result(PassTime)

    'time out fail
    If PassTime > 65 Then
        Call CloseMP_AU69XX
        MPContFail = MPContFail + 1
        TestResult = "Bin3"
        MPTester.TestResultLab = "Bin3:MP Time out Fail"
        MPTester.Print "MP Time out Fail"
        SetSiteStatus (SiteUnknow)
        Call TestEnd
    End If
       
    'MP fail
    If AlcorMPMessage = WM_FT_MP_FAIL Then
        Call CloseMP_AU69XX
        MPContFail = MPContFail + 1
        TestResult = "Bin3"
        MPTester.TestResultLab = "Bin3:MP Function Fail"
        MPTester.Print "MP Function Fail"
        SetSiteStatus (SiteUnknow)
        Call TestEnd
    End If

    'unknow fail
    If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
        Call CloseMP_AU69XX
        MPContFail = MPContFail + 1
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
        MPTester.Print "MP UNKNOW Fail"
        SetSiteStatus (SiteUnknow)
        Call TestEnd
    End If
       
    ' mp pass
    If AlcorMPMessage = WM_FT_MP_PASS Then
        MPTester.TestResultLab = "MP PASS"
        MPContFail = 0
        MPTester.Print "MP PASS"
        SetSiteStatus (MPDone)
    End If
End Sub

Public Function STD_RW_TEST() As Long
Dim mMsg As MSG
Dim PassTime As Long
Dim OldTimer As Long
Dim LedCount As Integer
Dim rt2
Dim rtLockCount As Integer

Dim LightSituation

    Lon = False
    Loff = False
    LightSituation = 255

    winHwnd = FindWindow(vbNullString, CurDevicePar.FT_ToolTitle)
    
    If winHwnd = 0 Then
        Call LoadFTtool_AU69XX
    
        If U2_Pass = False Then
            rt2 = PostMessage(winHwnd, WM_FT_TEST_U2, 0&, 0&)
            Call MsecDelay(0.5)
            winHwnd = FindWindow(vbNullString, "UFD Test")
        ElseIf U2_Pass = True Then
            rt2 = PostMessage(winHwnd, WM_FT_TEST_U3, 0&, 0&)
            Call MsecDelay(0.5)
            winHwnd = FindWindow(vbNullString, "UFD Test")
        End If
    
        MPTester.Print "wait for RW Tester Ready"
        OldTimer = Timer
        AlcorMPMessage = 0
        Do
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
            
            PassTime = Timer - OldTimer
        Loop Until AlcorMPMessage = WM_FT_RW_READY Or PassTime > 5 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
        
        MPTester.Print "RW Ready Time="; PassTime
        
        If PassTime > 5 Then
            CloseFTtool_AU69XX
            TestResult = "Bin3"
            MPTester.TestResultLab = "Bin3:RW Ready Fail"
            SetSiteStatus (SiteUnknow)
            'Call TestEnd
            STD_RW_TEST = PassTime
            Exit Function
        End If
    Else
RETRY_LOCK:
        If U2_Pass = False Then
            rt2 = PostMessage(winHwnd, WM_FT_TEST_U2, 0&, 0&)
            Call MsecDelay(0.5)
            winHwnd = FindWindow(vbNullString, "UFD Test")
        ElseIf U2_Pass = True Then
            rt2 = PostMessage(winHwnd, WM_FT_TEST_U3, 0&, 0&)
            Call MsecDelay(0.5)
            winHwnd = FindWindow(vbNullString, "UFD Test")
        End If
    End If

    If MPFlag = 1 Then
        WaitAnotherSiteDone (MPDone)
        Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
        If U3_Test Then
            cardresult = DO_WritePort(card, Channel_P1A, &H1)   'Power OFF UNLoad Device
        Else
            cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'Power OFF UNLoad Device
        End If
        
        WaitDevOn ("vid_058f")
        
        Call MsecDelay(0.3)
        
        Call PowerSet2(1, CurDevicePar.SetStdV1, CurDevicePar.SetStdI1, 1, CurDevicePar.SetStdV2, CurDevicePar.SetStdI2, 1)
        
        If Dual_Flag Then
            cardresult = DO_WritePort(card, Channel_P1A, &HFE)  '1110 關掉ena
            Call MsecDelay(0.04)
            cardresult = DO_WritePort(card, Channel_P1A, &HFA)  '1010 ena-b enabled sel to 厚
        Else
            If InStr(ChipName, "87100") Then
                If U2_Pass = False Then
                    cardresult = DO_WritePort(card, Channel_P1A, &HF6)
                Else
                    cardresult = DO_WritePort(card, Channel_P1A, &HF5)
                    MsecDelay (0.05)
                    cardresult = DO_WritePort(card, Channel_P1A, &HF4)
                End If
            Else
                cardresult = DO_WritePort(card, Channel_P1A, &HFB)
            End If
        End If
        
        For LedCount = 1 To 200
            Call MsecDelay(0.1)
            cardresult = DO_ReadPort(card, Channel_P1B, LightSituation)
            If LightSituation = 255 Then
                Loff = True
            Else
                Lon = True
            End If

            If (Loff = True) And (Lon = True) Then
                Exit For
            End If

        Next LedCount
        
        SetSiteStatus (RunHV)
        WaitAnotherSiteDone (RunHV)
        MPTester.Print "RW Tester begin test........"
        
        If (CurDevicePar.ShortName = "AU6992") Then
            Call MsecDelay(2)         ' test, u can modify this delay time shorter

            WaitDevOn ("vid_058f")
                
            Call MsecDelay(2)
            Call StartFTtest_AU69XX
        Else
            Call StartFTtest_WaitDevReady_AU69XX(4#, True)
        End If
        MPFlag = 0
         
    Else
        Call PowerSet2(1, CurDevicePar.SetStdV1, CurDevicePar.SetStdI1, 1, CurDevicePar.SetStdV2, CurDevicePar.SetStdI2, 1)
        
        If Dual_Flag Then
            cardresult = DO_WritePort(card, Channel_P1A, &HFE)
            Call MsecDelay(0.04)
            cardresult = DO_WritePort(card, Channel_P1A, &HFA)
        Else
            If InStr(ChipName, "87100") Then
                If U2_Pass = False Then
                    cardresult = DO_WritePort(card, Channel_P1A, &HF6)
                Else
                    cardresult = DO_WritePort(card, Channel_P1A, &HF5)
                    MsecDelay (0.05)
                    cardresult = DO_WritePort(card, Channel_P1A, &HF4)
                End If
            Else
                cardresult = DO_WritePort(card, Channel_P1A, &HFB)
            End If
        End If
        
        For LedCount = 1 To 200
            Call MsecDelay(0.1)
            cardresult = DO_ReadPort(card, Channel_P1B, LightSituation)
            If LightSituation = 255 Then
                Loff = True
            Else
                Lon = True
            End If

            If (Loff = True) And (Lon = True) Then
                Exit For
            End If

        Next LedCount
        
        SetSiteStatus (RunHV)
        WaitAnotherSiteDone (RunHV)
        MPTester.Print "RW Tester begin test........"
        
        If (CurDevicePar.ShortName = "AU6992") Then
            
            WaitDevOn ("vid_058f")
            
            Call MsecDelay(0.3)
            Call StartFTtest_AU69XX
        Else
            Call StartFTtest_WaitDevReady_AU69XX(4#)
            'Call MsecDelay(0.3)
        End If
       
    End If
             
    AlcorMPMessage = 0
    OldTimer = Timer
'    If (EQC_HV = True) And (EQC_LV = True) Then
'        Call StartFTtest_WaitDevReady_AU69XX(0.5)
'    End If
    
    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
        
        PassTime = Timer - OldTimer
       
    Loop Until AlcorMPMessage = WM_FT_RW_SPEED_FAIL _
            Or AlcorMPMessage = WM_FT_RW_RW_FAIL _
            Or AlcorMPMessage = WM_FT_RW_ROM_FAIL _
            Or AlcorMPMessage = WM_FT_RW_RAM_FAIL _
            Or AlcorMPMessage = WM_FT_RW_RW_PASS _
            Or AlcorMPMessage = WM_FT_RW_UNKNOW_FAIL _
            Or AlcorMPMessage = WM_FT_FLASH_NUM_FAIL _
            Or AlcorMPMessage = WM_FT_CHECK_CERBGPO_FAIL _
            Or AlcorMPMessage = WM_FT_CHECK_HW_CODE_FAIL _
            Or AlcorMPMessage = WM_FT_PHYREAD_FAIL _
            Or AlcorMPMessage = WM_FT_ECC_FAIL _
            Or AlcorMPMessage = WM_FT_NOFREEBLOCK_FAIL _
            Or AlcorMPMessage = WM_FT_LODECODE_FAIL _
            Or AlcorMPMessage = WM_FT_CHECK_WRITE_PROTECT_FAIL _
            Or AlcorMPMessage = WM_FT_MOVE_DATA_FAIL _
            Or AlcorMPMessage = WM_FT_NO_CARD_FAIL _
            Or AlcorMPMessage = WM_FT_LOCK_VOLUME_FAIL _
            Or AlcorMPMessage = WM_FT_UNLOCK_VOLUME_FAIL _
            Or AlcorMPMessage = WM_FT_TURNOFF_TUR_FAIL _
            Or AlcorMPMessage = WM_DEV_GET_HANDLE_FAIL _
            Or AlcorMPMessage = WM_DEV_GET_DIS_TUR_HANDLE_FAIL _
            Or AlcorMPMessage = WM_FT_TEST_DQS_FAIL Or AlcorMPMessage = WM_LC_FAIL Or AlcorMPMessage = WM_RC_FAIL Or PassTime > 10 _
            Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY


    If AlcorMPMessage = WM_FT_LOCK_VOLUME_FAIL Then
        'debug.print AlcorMPMessage
        rtLockCount = rtLockCount + 1
        If rtLockCount < 3 Then
            GoTo RETRY_LOCK
        End If
    End If
    
    MPTester.Print "RW work Time="; PassTime
    MPTester.MPText.Text = Hex(AlcorMPMessage)
    STD_RW_TEST = PassTime
    
    If (PassTime > 10) Or ((FailCloseAP) And (AlcorMPMessage <> WM_FT_RW_RW_PASS)) Then
        Close_FT_AP ("UFD Test")
        
        If (PassTime > 10) Then
            TestResult = "Bin3"
            MPTester.TestResultLab = "Bin3:RW Time Out Fail"
            'cardresult = DO_WritePort(card, Channel_P1A, &HFF)  ' power off
            SetSiteStatus (SiteUnknow)
            'Call TestEnd
            Exit Function
        End If
    End If
    
'    Call CloseFTtool_AU69XX
'
'    CurDevicePar.MP_ToolFileName = "AlcorMP_Test.exe"
'
'    Call LoadMP_AU69XX
'
'    'Call StartMP_AU69XX
'
'    AlcorMPMessage = 0
'    OldTimer = Timer
'
'    Do
'        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
'            AlcorMPMessage = mMsg.message
'            TranslateMessage mMsg
'            DispatchMessage mMsg
'        End If
'
'        PassTime = Timer - OldTimer
'
'    Loop Until AlcorMPMessage = WM_FT_MP_RECOGNIZE Or WM_FT_MP_UNRECOGNIZE Or PassTime > 10
'
'    Call CloseMP_AU69XX
'
'    Call LoadFTtool_AU69XX
    
    
End Function

Public Function AU87100_HLV_RW_TEST() As Long
Dim PassTime As Long
Dim OldTimer As Long
Dim mMsg As MSG
Dim rt2

    winHwnd = FindWindow(vbNullString, CurDevicePar.FT_ToolTitle)
    
    If winHwnd = 0 Then
             
        Call LoadFTtool_AU69XX
        
    End If
    
        If U2_Pass = False Then
            rt2 = PostMessage(winHwnd, WM_FT_TEST_U2, 0&, 0&)
            Call MsecDelay(0.5)
            winHwnd = FindWindow(vbNullString, "UFD Test")
        ElseIf U2_Pass = True Then
            rt2 = PostMessage(winHwnd, WM_FT_TEST_U3, 0&, 0&)
            Call MsecDelay(0.5)
            winHwnd = FindWindow(vbNullString, "UFD Test")
        End If
    
        MPTester.Print "wait for RW Tester Ready"
        OldTimer = Timer
        AlcorMPMessage = 0
        Do
            'DoEvents
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
            
            PassTime = Timer - OldTimer
        Loop Until AlcorMPMessage = WM_FT_RW_READY Or PassTime > 5 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
        MPTester.Print "RW Ready Time="; PassTime
        
       'HLV_RW_TEST = PassTime
        
        'GoTo T2
        If PassTime > 5 Then
            CloseFTtool_AU69XX
            TestResult = "Bin3"
            MPTester.TestResultLab = "Bin3:RW Ready Fail"
            SetSiteStatus (SiteUnknow)
            'Call TestEnd
            Exit Function
        End If
    'End If

    If MPFlag = 1 Then
            
        If (EQC_HV = False) And (EQC_LV = False) Then
            
            WaitAnotherSiteDone (MPDone)
            Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
            cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'Power OFF UNLoad Device
            
            WaitDevOFF ("vid_058f")
            
            Call MsecDelay(0.3)
            
            If Dir("D:\LABPC.PC") = "LABPC.PC" Then
                Call PowerSet2(1, "5.3", "0.2", 1, "5.3", "0.2", 1)
            Else
                Call PowerSet2(1, "0.0", CurDevicePar.SetHI1, 1, CurDevicePar.SetHV2, CurDevicePar.SetHI2, 1)
                Call MsecDelay(0.3)
                Call PowerSet2(1, CurDevicePar.SetHV1, CurDevicePar.SetHI1, 1, CurDevicePar.SetHV2, CurDevicePar.SetHI2, 1)
            End If
            
            If Dual_Flag Then
                cardresult = DO_WritePort(card, Channel_P1A, &HFE)  '1110 關掉ena
                Call MsecDelay(0.04)
                cardresult = DO_WritePort(card, Channel_P1A, &HFA)  '1010 ena-b enabled sel to 厚
            Else
                
                If InStr(ChipName, "87100") Then
                    If U2_Pass = False Then
                        cardresult = DO_WritePort(card, Channel_P1A, &HF6)
                    Else
                        cardresult = DO_WritePort(card, Channel_P1A, &HF5)
                        MsecDelay (0.05)
                        cardresult = DO_WritePort(card, Channel_P1A, &HF4)
                    End If
                Else
                    cardresult = DO_WritePort(card, Channel_P1A, &HFB)
                End If
                
            End If
            
            SetSiteStatus (RunHV)
            MPTester.Print "RW Tester begin test........"
            
            If (CurDevicePar.ShortName = "AU6992") Then
                
                WaitDevOn ("vid_058f")
                
                Call MsecDelay(0.3)
                Call StartFTtest_AU69XX
            Else
                Call StartFTtest_WaitDevReady_AU69XX(4#, True)
            End If
            
            SetSiteStatus (RunHV)
            
            EQC_HV = True
        End If
        MPFlag = 0
     
    Else
        If (EQC_HV = False) And (EQC_LV = False) Then
            
            If Dir("D:\LABPC.PC") = "LABPC.PC" Then
                Call PowerSet2(1, "5.3", "0.2", 1, "5.3", "0.2", 1)
            Else
                Call PowerSet2(1, "0.0", CurDevicePar.SetHI1, 1, CurDevicePar.SetHV2, CurDevicePar.SetHI2, 1)
                Call MsecDelay(0.3)
                Call PowerSet2(1, CurDevicePar.SetHV1, CurDevicePar.SetHI1, 1, CurDevicePar.SetHV2, CurDevicePar.SetHI2, 1)
            End If
            
            If Dual_Flag Then
                cardresult = DO_WritePort(card, Channel_P1A, &HFE)
                Call MsecDelay(0.04)
                cardresult = DO_WritePort(card, Channel_P1A, &HFA)
            Else
                If InStr(ChipName, "87100") Then
                    If U2_Pass = False Then
                        cardresult = DO_WritePort(card, Channel_P1A, &HF6)
                    Else
                        cardresult = DO_WritePort(card, Channel_P1A, &HF5)
                        MsecDelay (0.05)
                        cardresult = DO_WritePort(card, Channel_P1A, &HF4)
                    End If
                Else
                    cardresult = DO_WritePort(card, Channel_P1A, &HFB)
                End If
            End If
            
            'WaitDevOn ("vid_058f")
            SetSiteStatus (RunHV)
            MPTester.Print "RW Tester begin test........"
            
            If (CurDevicePar.ShortName = "AU6992") Then
                
                WaitDevOn ("vid_058f")

                Call MsecDelay(0.3)
                Call StartFTtest_AU69XX
            Else
                Call StartFTtest_WaitDevReady_AU69XX(4#)
            End If
'
'            If (InStr(1, ChipName, "19") <> 0 Or InStr(1, ChipName, "98") <> 0 And InStr(1, ChipName, "61") <> 0) Then
'                Call MsecDelay(0.4)
'            Else
'                Call MsecDelay(0.2)
'            End If
'
'            If (CurDevicePar.ShortName = "AU6988") Then
'                Call MsecDelay(0.2)
'            End If
            
            SetSiteStatus (RunHV)
            
            EQC_HV = True
        End If
    End If
                
    'T2:
    
    AlcorMPMessage = 0
    'MPTester.Print "RW Tester begin test........"
    OldTimer = Timer
    If (EQC_HV = True) And (EQC_LV = True) Then
    
'        Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
'        cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'Power OFF UNLoad Device
'        WaitDevOFF ("vid_058f")
'        Call MsecDelay(0.3)
'
'        Call PowerSet2(1, "0.0", CurDevicePar.SetHI1, 1, CurDevicePar.SetHV2, CurDevicePar.SetHI2, 1)
'        Call MsecDelay(0.3)
'        Call PowerSet2(1, CurDevicePar.SetHV1, CurDevicePar.SetHI1, 1, CurDevicePar.SetHV2, CurDevicePar.SetHI2, 1)
'
'        If U2_Pass = False Then
'            cardresult = DO_WritePort(card, Channel_P1A, &HF6)
'        Else
'            cardresult = DO_WritePort(card, Channel_P1A, &HF5)
'            MsecDelay (0.05)
'            cardresult = DO_WritePort(card, Channel_P1A, &HF4)
'        End If
    
        Call StartFTtest_WaitDevReady_AU69XX(3#, False)
    End If
    
    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
        
        PassTime = Timer - OldTimer
       
    Loop Until AlcorMPMessage = WM_FT_RW_SPEED_FAIL _
            Or AlcorMPMessage = WM_FT_RW_RW_FAIL _
            Or AlcorMPMessage = WM_FT_RW_ROM_FAIL _
            Or AlcorMPMessage = WM_FT_RW_RAM_FAIL _
            Or AlcorMPMessage = WM_FT_RW_RW_PASS _
            Or AlcorMPMessage = WM_FT_RW_UNKNOW_FAIL _
            Or AlcorMPMessage = WM_FT_FLASH_NUM_FAIL _
            Or AlcorMPMessage = WM_FT_CHECK_CERBGPO_FAIL _
            Or AlcorMPMessage = WM_FT_CHECK_HW_CODE_FAIL _
            Or AlcorMPMessage = WM_FT_PHYREAD_FAIL _
            Or AlcorMPMessage = WM_FT_ECC_FAIL _
            Or AlcorMPMessage = WM_FT_NOFREEBLOCK_FAIL _
            Or AlcorMPMessage = WM_FT_LODECODE_FAIL _
            Or AlcorMPMessage = WM_FT_CHECK_WRITE_PROTECT_FAIL _
            Or AlcorMPMessage = WM_FT_NO_CARD_FAIL _
            Or AlcorMPMessage = WM_FT_LOCK_VOLUME_FAIL _
            Or AlcorMPMessage = WM_FT_UNLOCK_VOLUME_FAIL _
            Or AlcorMPMessage = WM_FT_TURNOFF_TUR_FAIL _
            Or AlcorMPMessage = WM_DEV_GET_HANDLE_FAIL _
            Or AlcorMPMessage = WM_DEV_GET_DIS_TUR_HANDLE_FAIL _
            Or AlcorMPMessage = WM_FT_TEST_DQS_FAIL Or AlcorMPMessage = WM_LC_FAIL Or AlcorMPMessage = WM_RC_FAIL Or PassTime > 10 _
            Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
    MPTester.Print "RW work Time="; PassTime
    MPTester.MPText.Text = Hex(AlcorMPMessage)
    'HLV_RW_TEST = PassTime
    
    If (PassTime > 10) Or ((FailCloseAP) And (AlcorMPMessage <> WM_FT_RW_RW_PASS)) Then
        Close_FT_AP ("UFD Test")
        
        If (PassTime > 8) Then
            TestResult = "Bin3"
            MPTester.TestResultLab = "Bin3:RW Time Out Fail"
            SetSiteStatus (SiteUnknow)
            'AlcorMPMessage = WM_FT_RW_SPEED_FAIL
            Exit Function
        End If
    End If
    

End Function

Public Function HLV_RW_TEST() As Long
Dim PassTime As Long
Dim OldTimer As Long
Dim mMsg As MSG

    winHwnd = FindWindow(vbNullString, CurDevicePar.FT_ToolTitle)
    
    If winHwnd = 0 Then
             
        Call LoadFTtool_AU69XX
    
        MPTester.Print "wait for RW Tester Ready"
        OldTimer = Timer
        AlcorMPMessage = 0
        Do
            'DoEvents
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
            
            PassTime = Timer - OldTimer
        Loop Until AlcorMPMessage = WM_FT_RW_READY Or PassTime > 5 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
        MPTester.Print "RW Ready Time="; PassTime
        
        HLV_RW_TEST = PassTime
        
        'GoTo T2
        If PassTime > 5 Then
            CloseFTtool_AU69XX
            TestResult = "Bin3"
            MPTester.TestResultLab = "Bin3:RW Ready Fail"
            SetSiteStatus (SiteUnknow)
            'Call TestEnd
            Exit Function
        End If
    End If


    If MPFlag = 1 Then
            
        If (EQC_HV = False) And (EQC_LV = False) Then
            
            WaitAnotherSiteDone (MPDone)
            Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
            cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'Power OFF UNLoad Device

            WaitDevOFF ("vid_058f")

            Call MsecDelay(0.3)
            
            If Dir("D:\LABPC.PC") = "LABPC.PC" Then
                Call PowerSet2(1, "5.3", "0.2", 1, "5.3", "0.2", 1)
            Else
                Call PowerSet2(1, "0.0", CurDevicePar.SetHI1, 1, CurDevicePar.SetHV2, CurDevicePar.SetHI2, 1)
                Call MsecDelay(0.3)
                Call PowerSet2(1, CurDevicePar.SetHV1, CurDevicePar.SetHI1, 1, CurDevicePar.SetHV2, CurDevicePar.SetHI2, 1)
            End If
            
            If Dual_Flag Then
            cardresult = DO_WritePort(card, Channel_P1A, &HFE)  '1110 關掉ena
            Call MsecDelay(0.04)
            cardresult = DO_WritePort(card, Channel_P1A, &HFA)  '1010 ena-b enabled sel to 厚
            Else
                cardresult = DO_WritePort(card, Channel_P1A, &HFB)
            End If
            
            SetSiteStatus (RunHV)
            MPTester.Print "RW Tester begin test........"
            
            If (CurDevicePar.ShortName = "AU6992") Then
                
                WaitDevOFF ("vid_058f")

                Call MsecDelay(0.3)
                Call StartFTtest_AU69XX
            Else
                Call StartFTtest_WaitDevReady_AU69XX(6#, True)
            End If
            
            
'            If (InStr(1, ChipName, "19") <> 0 Or InStr(1, ChipName, "98") <> 0 And InStr(1, ChipName, "61") <> 0) Then
'                Call MsecDelay(0.4)
'            Else
'                Call MsecDelay(0.2)
'            End If
'
'            If (CurDevicePar.ShortName = "AU6988") Then
'                Call MsecDelay(0.2)
'            End If
            
            SetSiteStatus (RunHV)
            
            EQC_HV = True
        End If
        MPFlag = 0
     
    Else
        If (EQC_HV = False) Or (EQC_LV = False) Then
            
            If Dir("D:\LABPC.PC") = "LABPC.PC" Then
                Call PowerSet2(1, "5.3", "0.2", 1, "5.3", "0.2", 1)
            Else
                Call PowerSet2(1, "0.0", CurDevicePar.SetHI1, 1, CurDevicePar.SetHV2, CurDevicePar.SetHI2, 1)
                Call MsecDelay(0.3)
                Call PowerSet2(1, CurDevicePar.SetHV1, CurDevicePar.SetHI1, 1, CurDevicePar.SetHV2, CurDevicePar.SetHI2, 1)
            End If
            
            If Dual_Flag Then
                cardresult = DO_WritePort(card, Channel_P1A, &HFE)
                Call MsecDelay(0.04)
                cardresult = DO_WritePort(card, Channel_P1A, &HFA)
            Else
                cardresult = DO_WritePort(card, Channel_P1A, &HFB)
            End If
            
            SetSiteStatus (RunHV)
            MPTester.Print "RW Tester begin test........"
        
            If (CurDevicePar.ShortName = "AU6992") Then
    
                WaitDevOn ("vid_058f")
                
                Call MsecDelay(0.3)
                Call StartFTtest_AU69XX
            Else
                Call StartFTtest_WaitDevReady_AU69XX(10#)
            End If
'
'            If (InStr(1, ChipName, "19") <> 0 Or InStr(1, ChipName, "98") <> 0 And InStr(1, ChipName, "61") <> 0) Then
'                Call MsecDelay(0.4)
'            Else
'                Call MsecDelay(0.2)
'            End If
'
'            If (CurDevicePar.ShortName = "AU6988") Then
'                Call MsecDelay(0.2)
'            End If
            
            SetSiteStatus (RunHV)
            
            EQC_HV = True
        End If
    End If
                
    'T2:
    
    AlcorMPMessage = 0
    'MPTester.Print "RW Tester begin test........"
    OldTimer = Timer
    If (EQC_HV = True) And (EQC_LV = True) Then
        Call StartFTtest_WaitDevReady_AU69XX(0.5)
    End If
    
    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
        
        PassTime = Timer - OldTimer
       
    Loop Until AlcorMPMessage = WM_FT_RW_SPEED_FAIL _
            Or AlcorMPMessage = WM_FT_RW_RW_FAIL _
            Or AlcorMPMessage = WM_FT_RW_ROM_FAIL _
            Or AlcorMPMessage = WM_FT_RW_RAM_FAIL _
            Or AlcorMPMessage = WM_FT_RW_RW_PASS _
            Or AlcorMPMessage = WM_FT_RW_UNKNOW_FAIL _
            Or AlcorMPMessage = WM_FT_FLASH_NUM_FAIL _
            Or AlcorMPMessage = WM_FT_CHECK_CERBGPO_FAIL _
            Or AlcorMPMessage = WM_FT_CHECK_HW_CODE_FAIL _
            Or AlcorMPMessage = WM_FT_PHYREAD_FAIL _
            Or AlcorMPMessage = WM_FT_ECC_FAIL _
            Or AlcorMPMessage = WM_FT_NOFREEBLOCK_FAIL _
            Or AlcorMPMessage = WM_FT_LODECODE_FAIL _
            Or AlcorMPMessage = WM_FT_CHECK_WRITE_PROTECT_FAIL _
            Or AlcorMPMessage = WM_FT_NO_CARD_FAIL _
            Or AlcorMPMessage = WM_FT_LOCK_VOLUME_FAIL _
            Or AlcorMPMessage = WM_FT_UNLOCK_VOLUME_FAIL _
            Or AlcorMPMessage = WM_FT_TURNOFF_TUR_FAIL _
            Or AlcorMPMessage = WM_DEV_GET_HANDLE_FAIL _
            Or AlcorMPMessage = WM_DEV_GET_DIS_TUR_HANDLE_FAIL _
            Or AlcorMPMessage = WM_FT_TEST_DQS_FAIL Or AlcorMPMessage = WM_LC_FAIL Or AlcorMPMessage = WM_RC_FAIL Or PassTime > 10 _
            Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
    MPTester.Print "RW work Time="; PassTime
    MPTester.MPText.Text = Hex(AlcorMPMessage)
    HLV_RW_TEST = PassTime
    
    If (PassTime > 10) Or ((FailCloseAP) And (AlcorMPMessage <> WM_FT_RW_RW_PASS)) Then
        Close_FT_AP ("UFD Test")
        
        If (PassTime > 8) Then
            TestResult = "Bin3"
            MPTester.TestResultLab = "Bin3:RW Time Out Fail"
            SetSiteStatus (SiteUnknow)
            'AlcorMPMessage = WM_FT_RW_SPEED_FAIL
            Exit Function
        End If
    End If
    

End Function

Public Sub Check_STD_RW_Result(PassTime)
Dim LightOn
Dim LedCount As Byte
                    
    Select Case AlcorMPMessage
        Case WM_FT_RW_UNKNOW_FAIL
            TestResult = "Bin2"
            MPTester.TestResultLab = "Bin2:UnKnow Fail"
            ContFail = ContFail + 1
            
        Case WM_FT_FLASH_NUM_FAIL
            TestResult = "Bin2"
            MPTester.TestResultLab = "Bin2:Flash Number Fail"
            ContFail = ContFail + 1
        
        Case WM_FT_CHECK_HW_CODE_FAIL
            TestResult = "Bin5"
            MPTester.TestResultLab = "Bin5:HW-ID Fail"
            ContFail = ContFail + 1
        
'        Case WM_FT_TESTUNITREADY_FAIL
'            TestResult = "Bin2"
'            MPTester.TestResultLab = "Bin2:TestUnitReady Fail"
'            ContFail = ContFail + 1
        
        Case WM_FT_RW_SPEED_FAIL
            TestResult = "Bin3"
            MPTester.TestResultLab = "Bin3:SPEED Error "
            ContFail = ContFail + 1
             
        Case WM_FT_RW_RW_FAIL
            TestResult = "Bin3"
            MPTester.TestResultLab = "Bin3:RW FAIL "
            ContFail = ContFail + 1
        
        Case WM_FT_CHECK_CERBGPO_FAIL
            TestResult = "Bin3"
            MPTester.TestResultLab = "Bin3:GPO/RB FAIL "
            ContFail = ContFail + 1
        
        Case WM_FT_CHECK_WRITE_PROTECT_FAIL
            TestResult = "Bin3"
            MPTester.TestResultLab = "Bin3:W/P FAIL "
            ContFail = ContFail + 1
             
        Case WM_FT_NO_CARD_FAIL
            TestResult = "Bin4"
            MPTester.TestResultLab = "Bin4:NoCard FAIL "
            ContFail = ContFail + 1
        
        Case WM_FT_RW_ROM_FAIL
            TestResult = "Bin4"
            MPTester.TestResultLab = "Bin4:ROM FAIL "
            ContFail = ContFail + 1
              
        Case WM_FT_PHYREAD_FAIL
            TestResult = "Bin4"
            MPTester.TestResultLab = "Bin4:PHY Read FAIL "
            ContFail = ContFail + 1
              
        Case WM_FT_RW_RAM_FAIL
            TestResult = "Bin4"
            MPTester.TestResultLab = "Bin4:RAM FAIL "
            ContFail = ContFail + 1
               
        Case WM_FT_NOFREEBLOCK_FAIL
            TestResult = "Bin4"
            MPTester.TestResultLab = "Bin4:FreeBlock FAIL "
            ContFail = ContFail + 1
        
        Case WM_FT_LODECODE_FAIL
            TestResult = "Bin4"
            MPTester.TestResultLab = "Bin4:LoadCode FAIL "
            ContFail = ContFail + 1
        
'        Case WM_FT_RELOADCODE_FAIL
'            TestResult = "Bin4"
'            MPTester.TestResultLab = "Bin4:ReLoadCode FAIL "
'            ContFail = ContFail + 1
        
        Case WM_FT_ECC_FAIL
            TestResult = "Bin5"
            MPTester.TestResultLab = "Bin5:ECC FAIL "
            ContFail = ContFail + 1
            
        Case WM_FT_MOVE_DATA_FAIL
            TestResult = "Bin5"
            MPTester.TestResultLab = "Bin5:MOVE DATA FAIL "
            ContFail = ContFail + 1
            
        Case WM_DEV_GET_HANDLE_FAIL
            TestResult = "Bin2"
            MPTester.TestResultLab = "Bin2:Get Handle FAIL "
            ContFail = ContFail + 1
            
        Case WM_DEV_GET_DIS_TUR_HANDLE_FAIL
            TestResult = "Bin2"
            MPTester.TestResultLab = "Bin2:TURN ON TUR FAIL "
            ContFail = ContFail + 1
        
        Case WM_FT_LOCK_VOLUME_FAIL
           TestResult = "Bin2"
           MPTester.TestResultLab = "Bin2:LOCK VOLUME FAIL "
           ContFail = ContFail + 1
        
        Case WM_FT_UNLOCK_VOLUME_FAIL
           TestResult = "Bin2"
           MPTester.TestResultLab = "Bin2:UNLOCK VOLUME FAIL "
           ContFail = ContFail + 1
           
        Case WM_FT_TURNOFF_TUR_FAIL
           TestResult = "Bin2"
           MPTester.TestResultLab = "Bin2:TURN OFF TUR FAIL "
           ContFail = ContFail + 1
        
        Case WM_FT_TEST_DQS_FAIL
            TestResult = "Bin5"
            MPTester.TestResultLab = "Bin5:DQS check fail"
            ContFail = ContFail + 1
            
        Case WM_LC_FAIL
            TestResult = "Bin5"
            MPTester.TestResultLab = "Bin5: LC value fail"
            ContFail = ContFail + 1
            
        Case WM_RC_FAIL
            If InStr(ChipName, "2B") <> 0 Then   ' for test RC with record and binning
                TestResult = "Bin5"
                MPTester.TestResultLab = "Bin5: RC value fail"
                ContFail = ContFail + 1
            Else                                 ' for test RC with record but no binning
                If CheckLED_Flag Then
                    For LedCount = 1 To 20
                        Call MsecDelay(0.1)
                        cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
                        Debug.Print LightOn
                        If LightOn = &HEF Or LightOn = &HCF Or LightOn = 223 Or LightOn = 254 Then
                            Exit For
                        End If
                    Next LedCount
    
                    MPTester.Print "light="; LightOn
    
                    If ((LightOn = &HEF Or LightOn = &HCF Or LightOn = 223 Or LightOn = 254) And (Loff = True)) Then
                        MPTester.TestResultLab = "PASS "
                        TestResult = "PASS"
                        ContFail = 0
    
                        If U2_Pass = False Then U2_Pass = True
    
                    Else
                        TestResult = "Bin3"
                        MPTester.TestResultLab = "Bin3:LED FAIL "
                        ContFail = ContFail + 1
    
                        ' test use
                        'If U2_Pass = False Then U2_Pass = True
    
                    End If
                Else
                    MPTester.TestResultLab = "PASS "
                    TestResult = "PASS"
                    ContFail = 0
                    
                    'test use
                    If U2_Pass = False Then U2_Pass = True
                    
                End If
            End If
        

                    
        Case WM_FT_MP_RECOGNIZE
            TestResult = "PASS"
            MPTester.TestResultLab = "PASS"
            ContFail = 0
            
        Case WM_FT_MP_UNRECOGNIZE
            TestResult = "Bin2"
            MPTester.TestResultLab = "Bin2: MP UNRECOGNIZE"
            ContFail = ContFail + 1
                    
        Case WM_FT_RW_RW_PASS
            If CheckLED_Flag Then
                For LedCount = 1 To 20
                    Call MsecDelay(0.1)
                    cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
                    Debug.Print LightOn
                    If LightOn = &HEF Or LightOn = &HCF Or LightOn = 223 Or LightOn = 254 Then
                        Exit For
                    End If
                Next LedCount

                MPTester.Print "light="; LightOn

                If ((LightOn = &HEF Or LightOn = &HCF Or LightOn = 223 Or LightOn = 254) And (Loff = True)) Then
                    MPTester.TestResultLab = "PASS "
                    TestResult = "PASS"
                    ContFail = 0

                    If U2_Pass = False Then U2_Pass = True

                Else
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:LED FAIL "
                    ContFail = ContFail + 1

                    ' test use
                    'If U2_Pass = False Then U2_Pass = True

                End If
            Else
                MPTester.TestResultLab = "PASS "
                TestResult = "PASS"
                ContFail = 0
                
                'test use
                If U2_Pass = False Then U2_Pass = True
                
            End If
            
        Case Else
            TestResult = "Bin2"
            MPTester.TestResultLab = "Bin2:Undefine Fail"
            ContFail = ContFail + 1
    End Select
    
End Sub

Public Sub TestEnd()
    If InStr(ChipName, "87100") Then
        cardresult = DO_WritePort(card, Channel_P1A, &H1)
    Else
        cardresult = DO_WritePort(card, Channel_P1A, &HFF)
    End If
    
    Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
    SetSiteStatus (SiteUnknow)
    
    WaitDevOFF ("vid_058f")
    
    If (CurDevicePar.ShortName = "AU6991") Then
        Call MsecDelay(0.5)
    End If
End Sub

Public Sub AU87100HV_LVTestSub()

Dim OldTimer
Dim PassTime
Dim mMsg As MSG
Dim tmpName As String
Dim TimerCounter As Integer
Dim TmpString As String
Dim MP_Retry As Byte
Dim RWDone As Boolean
Dim refreshMSG

    'add unload driver function
    If PCI7248InitFinish = 0 Then
       PCI7248Exist
    End If
    
    MP_Retry = 0
    MPTester.TestResultLab = ""
    HV_Result = ""
    LV_Result = ""
    EQC_HV = False
    EQC_LV = False
    
    If Len(ChipName) = 15 Then
        tmpName = Left(ChipName, 10)
    Else
        tmpName = Left(ChipName, 9)
    End If
     
    MPTester.TestResultLab = ""
    ChDir App.Path & CurDevicePar.DeviceFolder
    
    NewChipFlag = 0
    '===============================================================
    ' Fail location initial
    '===============================================================
    If OldChipName <> ChipName Then
        Call Fail_Location_Initial(tmpName)
    End If
    
    OldChipName = ChipName
    SetSiteStatus (SiteReady)
    
    '==============================================================
    ' when begin RW Test, must clear MP program
    '===============================================================
    Call CloseMP_AU69XX
    
    MPTester.Print "ContFail="; ContFail
    MPTester.Print "MPContFail="; MPContFail
     
    '====================================
    '  Fix Card
    '====================================
    If InStr(ChipName, "87100") Then
        ContFail = 1
    End If
    
    If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag) Then
        If MPTester.NoMP.Value = 1 Then
            If (NewChipFlag = 0) And (MPTester.Check1.Value = 0) Then  ' force condition
                'Call HLV_RW_TEST(PassTime)
                RWDone = True
            End If
        End If
        
        If RWDone = False Then
            If MPTester.ResetMPFailCounter.Value = 1 Then
                ContFail = 0
            End If
            
            '===============================================================
            ' when begin MP, must close RW program
            '===============================================================
            MPFlag = 1
            U2_Pass = False
            Call CloseFTtool_AU69XX
            
            'power off
            If InStr(ChipName, "87100") Then
                cardresult = DO_WritePort(card, Channel_P1A, &H1)
            Else
                cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'close power
            End If
            Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)   ' close power to disable chip
            Call MsecDelay(0.5)  ' power for load MPDriver
            MPTester.Print "wait for MP Ready"
            Call LoadMP_AU69XX
        
            OldTimer = Timer
            AlcorMPMessage = 0
            Do
                'DoEvents
                If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                    AlcorMPMessage = mMsg.message
                    TranslateMessage mMsg
                    DispatchMessage mMsg
                End If
                PassTime = Timer - OldTimer
            
            Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 30 _
                    Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
        
            MPTester.Print "Ready Time="; PassTime
            
            '====================================================
            '  handle MP load time out, the FAIL will be Bin3
            '====================================================
            If PassTime > 30 Then
                Call CloseMP_AU69XX
                MPTester.TestResultLab = "Bin3:MP Ready Fail"
                TestResult = "Bin3"
                MPTester.Print "MP Ready Fail"
                SetSiteStatus (SiteUnknow)
                Call TestEnd
                Exit Sub
            End If
                
            '====================================================
            '  MP begin
            '====================================================
            
            If AlcorMPMessage = WM_FT_MP_START Then
                
                SetSiteStatus (RunMP)
                WaitAnotherSiteDone (RunMP)
                
                cardresult = DO_WritePort(card, Channel_P1A, &HFB)
                
                If Dir("D:\LABPC.PC") = "LABPC.PC" Then
                    Call PowerSet2(1, "5.0", "0.5", 1, "5.0", "0.5", 1)
                Else
                    Call PowerSet2(1, "0.0", "0.5", 1, CurDevicePar.SetHLVStd2, "0.5", 1)
                    Call MsecDelay(0.5)
                    Call PowerSet2(1, CurDevicePar.SetHLVStd1, "0.5", 1, CurDevicePar.SetHLVStd2, "0.5", 1)
                End If
                
                If InStr(ChipName, "87100") Then
                    If U2_Pass = False Then
                        cardresult = DO_WritePort(card, Channel_P1A, &HF6)
                    Else
                        cardresult = DO_WritePort(card, Channel_P1A, &HF5)
                        MsecDelay (0.05)
                        cardresult = DO_WritePort(card, Channel_P1A, &HF4)
                    End If
                Else
                    cardresult = DO_WritePort(card, Channel_P1A, &HFB)
                End If
                
                Call MsecDelay(0.5)
                
                Do
                    DoEvents
                    Call MsecDelay(0.1)
                    TimerCounter = TimerCounter + 1
                    TmpString = GetDeviceName("vid")
                Loop While (TmpString = "") And (TimerCounter < 100)
                     
                winHwnd = FindWindow(vbNullString, CurDevicePar.MP_WorkTitle)
                Call MsecDelay(0.1)
                refreshMSG = PostMessage(winHwnd, WM_FT_MP_REFRESH, 0&, 0&)
                    
                If TmpString = "" Then   ' can not find device after 15 s
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail when enter MP"
                    SetSiteStatus (SiteUnknow)
                    Call TestEnd
                    Exit Sub
                End If
                     
                Call MsecDelay(2.5)
                       
                MPTester.Print " MP Begin....."
                     
                Call StartMP_AU69XX
                OldTimer = Timer
                AlcorMPMessage = 0
                ReMP_Flag = 0
                
                Call MsecDelay(2.5)
                
                Do
                    'DoEvents
                    If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                        AlcorMPMessage = mMsg.message
                        TranslateMessage mMsg
                        DispatchMessage mMsg
                        
                        If (AlcorMPMessage = WM_FT_MP_FAIL) And (MP_Retry < 3) Then
                            'ReMP_Flag = 1
                            AlcorMPMessage = 1
                            
                            If InStr(ChipName, "87100") Then
                                cardresult = DO_WritePort(card, Channel_P1A, &H1)   'close power
                            Else
                                cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'close power
                            End If
                            
                            If Dual_Flag Then
                                cardresult = DO_WritePort(card, Channel_P1A, &HFD)
                            Else
                                If InStr(ChipName, "87100") Then
                                    If U2_Pass = False Then
                                        cardresult = DO_WritePort(card, Channel_P1A, &HF6)
                                    Else
                                        cardresult = DO_WritePort(card, Channel_P1A, &HF5)
                                        MsecDelay (0.05)
                                        cardresult = DO_WritePort(card, Channel_P1A, &HF4)
                                    End If
                                Else
                                    cardresult = DO_WritePort(card, Channel_P1A, &HFB)
                                End If
                            End If
                            
                            Call MsecDelay(2.2)
                            Call RefreshMP_AU69XX
                            Call MsecDelay(0.5)
                            Call StartMP_AU69XX
                            MP_Retry = MP_Retry + 1
                        End If
                            
                    End If
                    
                    PassTime = Timer - OldTimer
                
                Loop Until AlcorMPMessage = WM_FT_MP_PASS _
                        Or AlcorMPMessage = WM_FT_MP_FAIL _
                        Or AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL _
                        Or PassTime > 65 _
                        Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                
                MPTester.Print "MP work time="; PassTime
                MPTester.MPText.Text = Hex(AlcorMPMessage)

                '=========================================
                '    Close MP program
                '=========================================
                
                Call CloseMP_AU69XX
                KillProcess ("AlcorMP.exe")
                Call UnloadDriver
                
                '================================
                '  Handle MP work time out error
                '================================
                'Call Check_MP_Result(PassTime)
                'time out fail
                If PassTime > 65 Then
                    Call CloseMP_AU69XX
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    SetSiteStatus (SiteUnknow)
                    Call TestEnd
                    Exit Sub
                End If
                   
                'MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL Then
                    Call CloseMP_AU69XX
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    SetSiteStatus (SiteUnknow)
                    Call TestEnd
                    Exit Sub
                End If
            
                'unknow fail
                If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                    Call CloseMP_AU69XX
                    MPContFail = MPContFail + 1
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                    MPTester.Print "MP UNKNOW Fail"
                    SetSiteStatus (SiteUnknow)
                    Call TestEnd
                    Exit Sub
                End If
                   
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_PASS Then
                    MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPTester.Print "MP PASS"
                    SetSiteStatus (MPDone)
                End If
                
            End If
        End If
    End If
    
    Call AU87100_HLV_RW_TEST
    Call Check_HLV_RW_Result
    
    Call TestEnd
                   
End Sub

Public Sub AU69XXHV_LVTestSub()

Dim OldTimer
Dim PassTime
Dim mMsg As MSG
Dim tmpName As String
Dim TimerCounter As Integer
Dim TmpString As String
Dim MP_Retry As Byte
Dim RWDone As Boolean

    'add unload driver function
    If PCI7248InitFinish = 0 Then
       PCI7248Exist
    End If
    
    PLFlag = False  ' for AU6998-PL SOCKET
    
    KLFlag = False  ' for AU6927-KL
    
    MP_Retry = 0
    MPTester.TestResultLab = ""
    HV_Result = ""
    LV_Result = ""
    EQC_HV = False
    EQC_LV = False
    
    PLFlag = False  ' for AU6998-PL SOCKET
    
    If InStr(ChipName, "PL") Then
        PLFlag = True
    End If
    
    tmpName = Left(ChipName, 9)
     
    MPTester.TestResultLab = ""
    ChDir App.Path & CurDevicePar.DeviceFolder
    
    
    NewChipFlag = 0
    '===============================================================
    ' Fail location initial
    '===============================================================
    If OldChipName <> ChipName Then
        Call Fail_Location_Initial(tmpName)
        
        If Left(ChipName, 6) = "AU6921" Or Left(ChipName, 6) = "AU692A" Or Left(ChipName, 6) = "AU692A8" Or Left(ChipName, 6) = "AU692H" Or Left(ChipName, 6) = "AU692B" Then
            ' for MSL
            If (InStr(ChipName, "SLF")) Then
                FileCopy App.Path & CurDevicePar.DeviceFolder & "\62026UFDTest.exe", App.Path & CurDevicePar.DeviceFolder & "\UFDTest.exe"
            Else
                FileCopy App.Path & CurDevicePar.DeviceFolder & "\60026UFDTest.exe", App.Path & CurDevicePar.DeviceFolder & "\UFDTest.exe"
            End If
        End If
        
        If Left(ChipName, 6) = "AU6930" Then
            ' for MSL
            If (InStr(ChipName, "SLF")) Then
                FileCopy App.Path & CurDevicePar.DeviceFolder & "\62028UFDTest.exe", App.Path & CurDevicePar.DeviceFolder & "\UFDTest.exe"
            Else
                FileCopy App.Path & CurDevicePar.DeviceFolder & "\60028UFDTest.exe", App.Path & CurDevicePar.DeviceFolder & "\UFDTest.exe"
            End If
        End If
        
    End If
    
    OldChipName = ChipName
    SetSiteStatus (SiteReady)
    
    '==============================================================
    ' when begin RW Test, must clear MP program
    '===============================================================
    Call CloseMP_AU69XX
    
    MPTester.Print "ContFail="; ContFail
    MPTester.Print "MPContFail="; MPContFail
     
     
    '====================================
    '  Fix Card
    '====================================
    If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag) Then
        If MPTester.NoMP.Value = 1 Then
            If (NewChipFlag = 0) And (MPTester.Check1.Value = 0) Then  ' force condition
                'Call HLV_RW_TEST(PassTime)
                RWDone = True
            End If
        End If
        
        If RWDone = False Then
            If MPTester.ResetMPFailCounter.Value = 1 Then
                ContFail = 0
            End If
            
            '===============================================================
            ' when begin MP, must close RW program
            '===============================================================
               
            MPFlag = 1
            Call CloseFTtool_AU69XX
            
            'power on
            cardresult = DO_WritePort(card, Channel_P1A, &HFF)
            Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)   ' close power to disable chip
            Call MsecDelay(0.5)  ' power for load MPDriver
            MPTester.Print "wait for MP Ready"
            Call LoadMP_AU69XX
        
            OldTimer = Timer
            AlcorMPMessage = 0
            Do
                'DoEvents
                If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                    AlcorMPMessage = mMsg.message
                    TranslateMessage mMsg
                    DispatchMessage mMsg
                End If
                PassTime = Timer - OldTimer
            
            Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 30 _
                    Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
        
            MPTester.Print "Ready Time="; PassTime
            
            '====================================================
            '  handle MP load time out, the FAIL will be Bin3
            '====================================================
            If PassTime > 30 Then
                Call CloseMP_AU69XX
                MPTester.TestResultLab = "Bin3:MP Ready Fail"
                TestResult = "Bin3"
                MPTester.Print "MP Ready Fail"
                SetSiteStatus (SiteUnknow)
                Call TestEnd
                Exit Sub
            End If
                
            '====================================================
            '  MP begin
            '====================================================
            
            If AlcorMPMessage = WM_FT_MP_START Then
                
                SetSiteStatus (RunMP)
                WaitAnotherSiteDone (RunMP)
                
                If Dual_Flag Then
                    cardresult = DO_WritePort(card, Channel_P1A, &HFD)
                Else
                    cardresult = DO_WritePort(card, Channel_P1A, &HFB)
                End If
                
                If Dir("D:\LABPC.PC") = "LABPC.PC" Then
                    Call PowerSet2(1, "5.0", "0.5", 1, "5.0", "0.5", 1)
                Else
                    Call PowerSet2(1, "0.0", "0.5", 1, CurDevicePar.SetHLVStd2, "0.5", 1)
                    Call MsecDelay(0.5)
                    Call PowerSet2(1, CurDevicePar.SetHLVStd1, "0.5", 1, CurDevicePar.SetHLVStd2, "0.5", 1)
                End If
                
                Do
                    DoEvents
                    Call MsecDelay(0.1)
                    TimerCounter = TimerCounter + 1
                    TmpString = GetDeviceName("vid")
                Loop While (TmpString = "") And (TimerCounter < 100)
                     
                Call MsecDelay(0.3)
                    
                If TmpString = "" Then   ' can not find device after 15 s
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail when enter MP"
                    SetSiteStatus (SiteUnknow)
                    Call TestEnd
                    Exit Sub
                End If
                     
                Call MsecDelay(2.5)
                       
                MPTester.Print " MP Begin....."
                     
                Call StartMP_AU69XX
                OldTimer = Timer
                AlcorMPMessage = 0
                ReMP_Flag = 0
                
                Do
                    'DoEvents
                    If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                        AlcorMPMessage = mMsg.message
                        TranslateMessage mMsg
                        DispatchMessage mMsg
                        
                        If (AlcorMPMessage = WM_FT_MP_FAIL) And (MP_Retry < 3) Then
                            'ReMP_Flag = 1
                            AlcorMPMessage = 1
                            cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'close power
                            Call MsecDelay(0.3)
                            If Dual_Flag Then
                                cardresult = DO_WritePort(card, Channel_P1A, &HFD)
                            Else
                                cardresult = DO_WritePort(card, Channel_P1A, &HFB)
                            End If
                            Call MsecDelay(2.2)
                            Call RefreshMP_AU69XX
                            Call MsecDelay(0.5)
                            Call StartMP_AU69XX
                            MP_Retry = MP_Retry + 1
                        End If
                            
                    End If
                    
                    PassTime = Timer - OldTimer
                
                Loop Until AlcorMPMessage = WM_FT_MP_PASS _
                        Or AlcorMPMessage = WM_FT_MP_FAIL _
                        Or AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL _
                        Or PassTime > 65 _
                        Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                
                MPTester.Print "MP work time="; PassTime
                MPTester.MPText.Text = Hex(AlcorMPMessage)

                '=========================================
                '    Close MP program
                '=========================================
                
                Call CloseMP_AU69XX
                KillProcess ("AlcorMP.exe")
                Call UnloadDriver
                
                '================================
                '  Handle MP work time out error
                '================================
                'Call Check_MP_Result(PassTime)
                'time out fail
                If PassTime > 65 Then
                    Call CloseMP_AU69XX
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    SetSiteStatus (SiteUnknow)
                    Call TestEnd
                    Exit Sub
                End If
                   
                'MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL Then
                    Call CloseMP_AU69XX
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    SetSiteStatus (SiteUnknow)
                    Call TestEnd
                    Exit Sub
                End If
            
                'unknow fail
                If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                    Call CloseMP_AU69XX
                    MPContFail = MPContFail + 1
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                    MPTester.Print "MP UNKNOW Fail"
                    SetSiteStatus (SiteUnknow)
                    Call TestEnd
                    Exit Sub
                End If
                   
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_PASS Then
                    MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPTester.Print "MP PASS"
                    SetSiteStatus (MPDone)
                End If
                
            End If
        End If
    End If
    
    Call HLV_RW_TEST
    Call Check_HLV_RW_Result
    
    Call TestEnd
                   
End Sub

Public Sub Check_HLV_RW_Result()
Dim LightOn
Dim LedCount As Byte

    If (EQC_HV = True) And (EQC_LV = False) Then
           
        Select Case AlcorMPMessage
            
            Case WM_FT_RW_UNKNOW_FAIL
                TestResult = "Bin2"
                MPTester.TestResultLab = "HV: UnKnow Fail"
                
            Case WM_FT_FLASH_NUM_FAIL
                TestResult = "Bin2"
                MPTester.TestResultLab = "HV: Flash Number Fail"
            
            Case WM_FT_CHECK_HW_CODE_FAIL
                 TestResult = "Bin5"
                 MPTester.TestResultLab = "HV: HW-ID Fail"
            
'            Case WM_FT_TESTUNITREADY_FAIL
'                 TestResult = "Bin2"
'                 MPTester.TestResultLab = "HV: TestUnitReady Fail"
            
            Case WM_FT_RW_SPEED_FAIL
                 TestResult = "Bin3"
                 MPTester.TestResultLab = "HV: SPEED Error "
            
            Case WM_FT_RW_RW_FAIL
                 TestResult = "Bin3"
                 MPTester.TestResultLab = "HV: RW FAIL "
            
            Case WM_FT_CHECK_CERBGPO_FAIL
            
                TestResult = "Bin3"
                MPTester.TestResultLab = "HV: GPO/RB FAIL "
            
            Case WM_FT_CHECK_WRITE_PROTECT_FAIL
                 TestResult = "Bin3"
                 MPTester.TestResultLab = "HV: W/P FAIL "
                 
            Case WM_FT_NO_CARD_FAIL
                  TestResult = "Bin4"
                  MPTester.TestResultLab = "HV: NoCard FAIL "
            
            Case WM_FT_RW_ROM_FAIL
                TestResult = "Bin4"
                MPTester.TestResultLab = "HV: ROM FAIL "
            
            Case WM_FT_RW_RAM_FAIL
                TestResult = "Bin5"
                MPTester.TestResultLab = "HV: RAM FAIL "
            
            Case WM_FT_PHYREAD_FAIL
                TestResult = "Bin5"
                MPTester.TestResultLab = "HV: PHY Read FAIL"
                
            Case WM_FT_NOFREEBLOCK_FAIL
                TestResult = "Bin5"
                MPTester.TestResultLab = "HV: NOFREEBLOCK FAIL"
            
            Case WM_FT_LODECODE_FAIL
                TestResult = "Bin5"
                MPTester.TestResultLab = "HV: LODECODE FAIL"
            
            Case WM_FT_ECC_FAIL
                TestResult = "Bin5"
                MPTester.TestResultLab = "HV: ECC FAIL"
            
'            Case WM_FT_RELOADCODE_FAIL
'                TestResult = "Bin5"
'                MPTester.TestResultLab = "HV: RELOADCODE FAIL"
            
            Case WM_FT_MOVE_DATA_FAIL
                TestResult = "Bin5"
                MPTester.TestResultLab = "HV: MOVE DATA FAIL"
                
            Case WM_DEV_GET_HANDLE_FAIL
                TestResult = "Bin2"
                MPTester.TestResultLab = "HV:Get Handle FAIL "
            
            Case WM_DEV_GET_DIS_TUR_HANDLE_FAIL
                TestResult = "Bin2"
                MPTester.TestResultLab = "HV:TURN ON TUR FAIL "
        
            Case WM_FT_LOCK_VOLUME_FAIL
               TestResult = "Bin2"
               MPTester.TestResultLab = "HV:LOCK VOLUME FAIL "
        
            Case WM_FT_UNLOCK_VOLUME_FAIL
               TestResult = "Bin2"
               MPTester.TestResultLab = "HV:UNLOCK VOLUME FAIL "
           
            Case WM_FT_TEST_DQS_FAIL
                TestResult = "Bin5"
                MPTester.TestResultLab = "HV:Bin5:DQS fail"
            
            Case WM_LC_FAIL
                TestResult = "Bin5"
                MPTester.TestResultLab = "HV:Bin5: LC value fail"
                
            Case WM_RC_FAIL
                If InStr(ChipName, "2B") <> 0 Then   ' for test RC with record and binning
                    TestResult = "Bin5"
                    MPTester.TestResultLab = "HV:Bin5: RC value fail"
                Else                                 ' for test RC with record but no binning
                    If CheckLED_Flag Then
                        For LedCount = 1 To 20
                            Call MsecDelay(0.1)
                            cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
                            Debug.Print LightOn
                            If LightOn = &HEF Or LightOn = &HCF Or LightOn = 223 Or LightOn = 254 Then
                                Exit For
                            End If
                        Next LedCount
    
                        MPTester.Print "light="; LightOn
    
                        If (LightOn = &HEF Or LightOn = &HCF Or LightOn = 223 Or LightOn = 254) Then
                            MPTester.TestResultLab = "HV: PASS "
                            TestResult = "PASS"
                        Else
                            MPTester.TestResultLab = "HV: LED FAIL "
                            TestResult = "Bin3"
                        End If
                        
                    Else
                        MPTester.TestResultLab = "HV: PASS "
                        TestResult = "PASS"
                    End If
                End If

            Case WM_FT_TURNOFF_TUR_FAIL
               TestResult = "Bin2"
               MPTester.TestResultLab = "HV:TURN OFF TUR FAIL "
               
            Case WM_FT_RW_RW_PASS
                
                If CheckLED_Flag Then
                    For LedCount = 1 To 20
                        Call MsecDelay(0.1)
                        cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
                        Debug.Print LightOn
                        If LightOn = &HEF Or LightOn = &HCF Or LightOn = 223 Or LightOn = 254 Then
                            Exit For
                        End If
                    Next LedCount

                    MPTester.Print "light="; LightOn

                    If (LightOn = &HEF Or LightOn = &HCF Or LightOn = 223 Or LightOn = 254) Then
                        MPTester.TestResultLab = "HV: PASS "
                        TestResult = "PASS"
                    Else
                        MPTester.TestResultLab = "HV: LED FAIL "
                        TestResult = "Bin3"
                    End If
                    
                Else
                    MPTester.TestResultLab = "HV: PASS "
                    TestResult = "PASS"
                End If
                
            Case Else
                 TestResult = "Bin2"
                 MPTester.TestResultLab = "HV: Undefine Fail"
        
        End Select
    
        
        HV_Result = TestResult
        TestResult = ""
        
        SetSiteStatus (HVDone)
        WaitAnotherSiteDone (HVDone)
        
        Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
        cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'Power OFF UNLoad Device
        
        WaitDevOFF ("vid_058f")

        Call MsecDelay(0.1)
        If (CurDevicePar.ShortName = "AU6991") Then
            Call MsecDelay(0.5)
        End If
        
        SetSiteStatus (RunLV)
        WaitAnotherSiteDone (RunLV)
        
        If Dir("D:\LABPC.PC") = "LABPC.PC" Then
            Call PowerSet2(1, "4.8", "0.15", 1, "4.8", "0.15", 1)
        Else
            Call PowerSet2(1, "0.0", CurDevicePar.SetLI1, 1, CurDevicePar.SetLV2, CurDevicePar.SetLI2, 1)
            Call MsecDelay(0.3)
            Call PowerSet2(1, CurDevicePar.SetLV1, CurDevicePar.SetLI1, 1, CurDevicePar.SetLV2, CurDevicePar.SetLI2, 1)
        End If
        
        If Dual_Flag Then
            cardresult = DO_WritePort(card, Channel_P1A, &HFA)
        Else
            If InStr(ChipName, "87100") Then
                If U2_Pass = False Then
                    cardresult = DO_WritePort(card, Channel_P1A, &HF6)
                Else
                    cardresult = DO_WritePort(card, Channel_P1A, &HF5)
                    MsecDelay (0.05)
                    cardresult = DO_WritePort(card, Channel_P1A, &HF4)
                End If
            Else
                cardresult = DO_WritePort(card, Channel_P1A, &HFB)
            End If
        End If
        
        If MPFlag = 1 Then
            Call MsecDelay(4)
        End If

        WaitDevOn ("vid_058f")
        
        If (InStr(1, ChipName, "19") <> 0 Or InStr(1, ChipName, "98") <> 0 And InStr(1, ChipName, "61") <> 0) Then
            Call MsecDelay(0.4)
        Else
            Call MsecDelay(0.2)
        End If
        
        If (CurDevicePar.ShortName = "AU6988") Then
            Call MsecDelay(0.2)
        End If
        
        If InStr(ChipName, "AU87100") Then
            Call AU87100_HLV_RW_TEST
        Else
            Call HLV_RW_TEST
        End If
            
        EQC_LV = True
        Call Check_HLV_RW_Result
        
        
    ElseIf (EQC_HV = True) And (EQC_LV = True) Then
            
        Select Case AlcorMPMessage
    
        Case WM_FT_RW_UNKNOW_FAIL
            TestResult = "Bin2"
            MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: UnKnow Fail"
            
        Case WM_FT_FLASH_NUM_FAIL
            TestResult = "Bin2"
            MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: Flash Number Fail"
            
        Case WM_FT_CHECK_HW_CODE_FAIL
             TestResult = "Bin5"
             MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: HW-ID Fail"
            
'        Case WM_FT_TESTUNITREADY_FAIL
'             TestResult = "Bin2"
'             MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: TestUnitReady Fail"
             
        Case WM_FT_RW_SPEED_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: SPEED Error "
             
        Case WM_FT_RW_RW_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: RW FAIL "
             
        Case WM_FT_TEST_DQS_FAIL
             TestResult = "Bin5"
             MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: DQS fail "
        
        Case WM_LC_FAIL
            TestResult = "Bin5"
            MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: LC value fail"
            
        Case WM_RC_FAIL
            If InStr(ChipName, "2B") <> 0 Then   ' for test RC with record and binning
                TestResult = "Bin5"
                MPTester.TestResultLab = "HV:Bin5: RC value fail"
            Else                                 ' for test RC with record but no binning
                If CheckLED_Flag Then
                    For LedCount = 1 To 20
                        Call MsecDelay(0.1)
                        cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
    
                        If LightOn = &HEF Or LightOn = &HCF Or LightOn = 223 Or LightOn = 254 Then
                            Exit For
                        End If
                    Next LedCount
    
                    MPTester.Print "light="; LightOn
    
                    If (LightOn = &HEF Or LightOn = &HCF Or LightOn = 223 Or LightOn = 254) Then
                        MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: PASS"
                        TestResult = "PASS"
                    Else
                        MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: LED FAIL "
                        TestResult = "Bin3"
                    End If
                
                Else
                    MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: PASS"
                    TestResult = "PASS"
                End If
            End If
             
        Case WM_FT_CHECK_CERBGPO_FAIL
    
            TestResult = "Bin3"
            MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: GPO/RB FAIL "
            
        Case WM_FT_CHECK_WRITE_PROTECT_FAIL
            TestResult = "Bin3"
            MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: W/P FAIL "
             
        Case WM_FT_NO_CARD_FAIL
            TestResult = "Bin4"
            MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: NoCard FAIL "
            
        Case WM_FT_RW_ROM_FAIL
            TestResult = "Bin4"
            MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: ROM FAIL "
            
        Case WM_FT_RW_RAM_FAIL, WM_FT_PHYREAD_FAIL, WM_FT_ECC_FAIL, WM_FT_NOFREEBLOCK_FAIL, WM_FT_LODECODE_FAIL, WM_FT_MOVE_DATA_FAIL
            TestResult = "Bin4"
            MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: RAM FAIL "
        
        Case WM_DEV_GET_HANDLE_FAIL
            TestResult = "Bin2"
            MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV:Get Handle FAIL "
            
        Case WM_DEV_GET_DIS_TUR_HANDLE_FAIL
            TestResult = "Bin2"
            MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV:TURN ON TUR FAIL "
    
        Case WM_FT_LOCK_VOLUME_FAIL
           TestResult = "Bin2"
           MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV:LOCK VOLUME FAIL "
    
        Case WM_FT_UNLOCK_VOLUME_FAIL
           TestResult = "Bin2"
           MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV:UNLOCK VOLUME FAIL "
       
        Case WM_FT_TURNOFF_TUR_FAIL
           TestResult = "Bin2"
           MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV:TURN OFF TUR FAIL "
               
        Case WM_FT_RW_RW_PASS
            
            If CheckLED_Flag Then
                For LedCount = 1 To 20
                    Call MsecDelay(0.1)
                    cardresult = DO_ReadPort(card, Channel_P1B, LightOn)

                    If LightOn = &HEF Or LightOn = &HCF Or LightOn = 223 Or LightOn = 254 Then
                        Exit For
                    End If
                Next LedCount

                MPTester.Print "light="; LightOn

                If (LightOn = &HEF Or LightOn = &HCF Or LightOn = 223 Or LightOn = 254) Then
                    MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: PASS"
                    TestResult = "PASS"
                Else
                    MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: LED FAIL "
                    TestResult = "Bin3"
                End If
            
            Else
                MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: PASS"
                TestResult = "PASS"
            End If
        
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: Undefine Fail"
                   
        End Select
          
        LV_Result = TestResult
        TestResult = ""
        
        If (HV_Result = "Bin2") And (LV_Result = "Bin2") Then
            TestResult = "Bin2"
            ContFail = ContFail + 1
        ElseIf (HV_Result <> "PASS") And (LV_Result = "PASS") Then
            TestResult = "Bin3"
            ContFail = ContFail + 1
        ElseIf (HV_Result = "PASS") And (LV_Result <> "PASS") Then
            TestResult = "Bin4"
            ContFail = ContFail + 1
        ElseIf (HV_Result <> "PASS") And (LV_Result <> "PASS") Then
            TestResult = "Bin5"
            ContFail = ContFail + 1
        ElseIf (HV_Result = "PASS") And (LV_Result = "PASS") Then
        
            If U2_Pass = False Then
                U2_Pass = True
            End If
            
            TestResult = "PASS"
            ContFail = 0
        Else
            TestResult = "Bin2"
            ContFail = ContFail + 1
        End If
        
        SetSiteStatus (LVDone)
        WaitAnotherSiteDone (LVDone)
      
    End If
End Sub

Public Sub AU6919ST3TestSub()

'AU6919 FT2 + FT2

Dim OldTimer
Dim PassTime
Dim rt2
Dim LightOn
Dim mMsg As MSG
Dim LedCount As Byte
Dim HV_Result As String
Dim LV_Result As String
Dim tmpName As String
Dim TimerCounter As Integer
Dim TmpString As String
Dim MP_Retry As Byte

    'add unload driver function
    If PCI7248InitFinish = 0 Then
       PCI7248Exist
    End If
    
    MP_Retry = 0
    MPTester.TestResultLab = ""
    HV_Result = ""
    LV_Result = ""
    EQC_HV = False
    EQC_LV = False
    tmpName = Left(ChipName, 9)
     
    MPTester.TestResultLab = ""
    ChDir App.Path & CurDevicePar.DeviceFolder
    
    NewChipFlag = 0
    If OldChipName <> ChipName Then
        '===============================================================
        ' Fail location initial
        '===============================================================
        If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
            FileCopy App.Path & CurDevicePar.DeviceFolder & "\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
            Call MsecDelay(5)
        End If
    
        If (OldVer_Flag) Then
            FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\ROM.Hex", App.Path & CurDevicePar.DeviceFolder & "\ROM.Hex"
            FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\RAM.Bin", App.Path & CurDevicePar.DeviceFolder & "\RAM.Bin"
            FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\AlcorMP.ini", App.Path & CurDevicePar.DeviceFolder & "\AlcorMP.ini"
            FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\PE.bin", App.Path & CurDevicePar.DeviceFolder & "\PE.bin"
            FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\FT.ini", App.Path & CurDevicePar.DeviceFolder & "\FT.ini"
        Else
            FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\ROM.Hex", App.Path & CurDevicePar.DeviceFolder & "\ROM.Hex"
            FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\RAM.Bin", App.Path & CurDevicePar.DeviceFolder & "\RAM.Bin"
            FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\AlcorMP.ini", App.Path & CurDevicePar.DeviceFolder & "\AlcorMP.ini"
            FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\PE.bin", App.Path & CurDevicePar.DeviceFolder & "\PE.bin"
            
            If PLFlag = True Then
                FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\PL_FT.ini", App.Path & CurDevicePar.DeviceFolder & "\FT.ini"
            ElseIf KLFlag = True Then
                FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\KL_FT.ini", App.Path & CurDevicePar.DeviceFolder & "\FT.ini"
            Else
                FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\FT.ini", App.Path & CurDevicePar.DeviceFolder & "\FT.ini"
            End If
        End If
        
        'FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & CurDevicePar.ShortName & "\FT.ini", App.Path & "\FT.ini"
        NewChipFlag = 1 ' force MP
    End If
    
    OldChipName = ChipName
    SetSiteStatus (SiteReady)
    
    '==============================================================
    ' when begin RW Test, must clear MP program
    '===============================================================
    Call CloseMP_AU69XX
    
    MPTester.Print "ContFail="; ContFail
    MPTester.Print "MPContFail="; MPContFail
     
     
    '====================================
    '  Fix Card
    '====================================
    ' GoTo T1
    If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag) Then
        If MPTester.NoMP.Value = 1 Then
            If (NewChipFlag = 0) And (MPTester.Check1.Value = 0) Then  ' force condition
                GoTo RW_Test_Label
            End If
        End If
        
        If MPTester.ResetMPFailCounter.Value = 1 Then
            ContFail = 0
        End If
        
        '===============================================================
        ' when begin MP, must close RW program
        '===============================================================
           
        MPFlag = 1
     
        Call CloseFTtool_AU69XX
        
        'power on
        cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)   ' close power to disable chip
        Call MsecDelay(0.5)  ' power for load MPDriver
        MPTester.Print "wait for MP Ready"
        Call LoadMP_AU69XX
    
        OldTimer = Timer
        AlcorMPMessage = 0
        ''debug.print "begin"
        Do
            'DoEvents
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
            PassTime = Timer - OldTimer
        
        Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 30 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
    
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3
        '====================================================
        If PassTime > 30 Then
        
            Call CloseMP_AU69XX
            MPTester.TestResultLab = "Bin3:MP Ready Fail"
            TestResult = "Bin3"
            MPTester.Print "MP Ready Fail"
            SetSiteStatus (SiteUnknow)
            GoTo TestEnd
            'Exit Sub
        End If
            
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_MP_START Then
            
            SetSiteStatus (RunMP)
            WaitAnotherSiteDone (RunMP)
            
            If Dual_Flag Then
                cardresult = DO_WritePort(card, Channel_P1A, &HFD)
            Else
                cardresult = DO_WritePort(card, Channel_P1A, &HFB)
            End If
            
            If Dir("D:\LABPC.PC") = "LABPC.PC" Then
                Call PowerSet2(1, "5.0", "0.5", 1, "5.0", "0.5", 1)
            Else
                Call PowerSet2(1, CurDevicePar.SetHLVStd1, "0.5", 1, CurDevicePar.SetHLVStd2, "0.5", 1)
            End If
            
            Do
                DoEvents
                Call MsecDelay(0.1)
                TimerCounter = TimerCounter + 1
                TmpString = GetDeviceName("vid")
            Loop While (TmpString = "") And (TimerCounter < 100)
                 
            Call MsecDelay(0.3)
                
            If TmpString = "" Then   ' can not find device after 15 s
                TestResult = "Bin2"
                MPTester.TestResultLab = "Bin2:MP UNKNOW Fail when enter MP"
                GoTo TestEnd
            End If
                 
            Call MsecDelay(2.5)
                   
            MPTester.Print " MP Begin....."
                 
            Call StartMP_AU69XX
            OldTimer = Timer
            AlcorMPMessage = 0
            ReMP_Flag = 0
            
            Do
                'DoEvents
                If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                    AlcorMPMessage = mMsg.message
                    TranslateMessage mMsg
                    DispatchMessage mMsg
                    
                    If (AlcorMPMessage = WM_FT_MP_FAIL) And (MP_Retry < 3) Then
                        'ReMP_Flag = 1
                        AlcorMPMessage = 1
                        cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'close power
                        Call MsecDelay(0.3)
                        If Dual_Flag Then
                            cardresult = DO_WritePort(card, Channel_P1A, &HFD)
                        Else
                            cardresult = DO_WritePort(card, Channel_P1A, &HFB)
                        End If
                        Call MsecDelay(2.2)
                        Call RefreshMP_AU69XX
                        Call MsecDelay(0.5)
                        Call StartMP_AU69XX
                        MP_Retry = MP_Retry + 1
                    End If
                        
                End If
                
                PassTime = Timer - OldTimer
            
            Loop Until AlcorMPMessage = WM_FT_MP_PASS _
                    Or AlcorMPMessage = WM_FT_MP_FAIL _
                    Or AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL _
                    Or PassTime > 65 _
                    Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
            
            MPTester.Print "MP work time="; PassTime
            MPTester.MPText.Text = Hex(AlcorMPMessage)
            '================================================
            '  Handle MP work time out error
            '===============================================
            
            'time out fail
            If PassTime > 65 Then
                Call CloseMP_AU69XX
                MPContFail = MPContFail + 1
                TestResult = "Bin3"
                MPTester.TestResultLab = "Bin3:MP Time out Fail"
                MPTester.Print "MP Time out Fail"
                SetSiteStatus (SiteUnknow)
                GoTo TestEnd
            End If
            
                
            'MP fail
            If AlcorMPMessage = WM_FT_MP_FAIL Then
                Call CloseMP_AU69XX
                MPContFail = MPContFail + 1
                TestResult = "Bin3"
                MPTester.TestResultLab = "Bin3:MP Function Fail"
                MPTester.Print "MP Function Fail"
                SetSiteStatus (SiteUnknow)
                GoTo TestEnd
            End If
        
                    
            'unknow fail
            If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                Call CloseMP_AU69XX
                MPContFail = MPContFail + 1
                TestResult = "Bin2"
                MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                MPTester.Print "MP UNKNOW Fail"
                SetSiteStatus (SiteUnknow)
                GoTo TestEnd
            End If
             
                       
            ' mp pass
            If AlcorMPMessage = WM_FT_MP_PASS Then
                MPTester.TestResultLab = "MP PASS"
                MPContFail = 0
                MPTester.Print "MP PASS"
                SetSiteStatus (MPDone)
            End If
        End If
       
    End If
    
    
    '=========================================
    '    Close MP program
    '=========================================
    
    Call CloseMP_AU69XX
    KillProcess ("AlcorMP.exe")
    Call UnloadDriver
    
    '=========================================
    '    POWER on
    '=========================================
    
RW_Test_Label:
    
    winHwnd = FindWindow(vbNullString, CurDevicePar.FT_ToolTitle)
    
    If winHwnd = 0 Then
             
        Call LoadFTtool_AU69XX
    
        MPTester.Print "wait for RW Tester Ready"
        OldTimer = Timer
        AlcorMPMessage = 0
        Do
            'DoEvents
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
            
            PassTime = Timer - OldTimer
        Loop Until AlcorMPMessage = WM_FT_RW_READY Or PassTime > 5 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
        MPTester.Print "RW Ready Time="; PassTime
        
        
        'GoTo T2
        If PassTime > 5 Then
            CloseFTtool_AU69XX
            TestResult = "Bin3"
            MPTester.TestResultLab = "Bin3:RW Ready Fail"
            SetSiteStatus (SiteUnknow)
            GoTo TestEnd
            'Exit Sub
        End If
    End If
    
    If MPFlag = 1 Then
            
        If (EQC_HV = False) And (EQC_LV = False) Then
            
            WaitAnotherSiteDone (MPDone)
            Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
            cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'Power OFF UNLoad Device
            WaitDevOFF ("vid_058f")
            Call MsecDelay(0.3)
            
            If Dir("D:\LABPC.PC") = "LABPC.PC" Then
                Call PowerSet2(1, "5.3", "0.2", 1, "5.3", "0.2", 1)
            Else
                Call PowerSet2(1, CurDevicePar.SetHV1, CurDevicePar.SetHI1, 1, CurDevicePar.SetHV2, CurDevicePar.SetHI2, 1)
            End If
            
            If Dual_Flag Then
                cardresult = DO_WritePort(card, Channel_P1A, &HFE)
                Call MsecDelay(0.04)
                cardresult = DO_WritePort(card, Channel_P1A, &HFA)
            Else
                cardresult = DO_WritePort(card, Channel_P1A, &HFB)
            End If
            
            WaitDevOn ("vid_058f")
            Call MsecDelay(0.2)
            
            SetSiteStatus (RunHV)
            
            EQC_HV = True
        End If
        MPFlag = 0
     
    Else
        If (EQC_HV = False) And (EQC_LV = False) Then
            
            If Dir("D:\LABPC.PC") = "LABPC.PC" Then
                Call PowerSet2(1, "5.3", "0.2", 1, "5.3", "0.2", 1)
            Else
                Call PowerSet2(1, CurDevicePar.SetHV1, CurDevicePar.SetHI1, 1, CurDevicePar.SetHV2, CurDevicePar.SetHI2, 1)
            End If
            If Dual_Flag Then
                cardresult = DO_WritePort(card, Channel_P1A, &HFE)
                Call MsecDelay(0.04)
                cardresult = DO_WritePort(card, Channel_P1A, &HFA)
            Else
                cardresult = DO_WritePort(card, Channel_P1A, &HFB)
            End If
            
            WaitDevOn ("vid_058f")
            Call MsecDelay(0.2)
            
            SetSiteStatus (RunHV)
            
            EQC_HV = True
        End If
    End If
             
    'T2:
    OldTimer = Timer
    AlcorMPMessage = 0
    MPTester.Print "RW Tester begin test........"
    Call StartFTtest_AU69XX
    
    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
        
        PassTime = Timer - OldTimer
       
    Loop Until AlcorMPMessage = WM_FT_RW_SPEED_FAIL _
            Or AlcorMPMessage = WM_FT_RW_RW_FAIL _
            Or AlcorMPMessage = WM_FT_RW_ROM_FAIL _
            Or AlcorMPMessage = WM_FT_RW_RAM_FAIL _
            Or AlcorMPMessage = WM_FT_RW_RW_PASS _
            Or AlcorMPMessage = WM_FT_RW_UNKNOW_FAIL _
            Or AlcorMPMessage = WM_FT_FLASH_NUM_FAIL _
            Or AlcorMPMessage = WM_FT_CHECK_CERBGPO_FAIL _
            Or AlcorMPMessage = WM_FT_CHECK_HW_CODE_FAIL _
            Or AlcorMPMessage = WM_FT_PHYREAD_FAIL _
            Or AlcorMPMessage = WM_FT_ECC_FAIL _
            Or AlcorMPMessage = WM_FT_NOFREEBLOCK_FAIL _
            Or AlcorMPMessage = WM_FT_LODECODE_FAIL _
'            Or AlcorMPMessage = WM_FT_TESTUNITREADY_FAIL _
'            Or AlcorMPMessage = WM_FT_RELOADCODE_FAIL _
            Or AlcorMPMessage = WM_FT_CHECK_WRITE_PROTECT_FAIL _
            Or AlcorMPMessage = WM_FT_NO_CARD_FAIL _
            Or AlcorMPMessage = WM_FT_TEST_DQS_FAIL Or AlcorMPMessage = WM_LC_FAIL Or AlcorMPMessage = WM_RC_FAIL Or PassTime > 10 _
            Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
             
                    
                    
    MPTester.Print "RW work Time="; PassTime
    MPTester.MPText.Text = Hex(AlcorMPMessage)
            
    '===========================================================
    '  RW Time Out Fail
    '===========================================================
    
    If (PassTime > 10) Or ((FailCloseAP) And (AlcorMPMessage <> WM_FT_RW_RW_PASS)) Then
        Close_FT_AP ("UFD Test")
        
        If (PassTime > 10) Then
            TestResult = "Bin3"
            MPTester.TestResultLab = "Bin3:RW Time Out Fail"
            AlcorMPMessage = WM_FT_RW_SPEED_FAIL
        End If
    
    End If
         
                   
    If (EQC_HV = True) And (EQC_LV = False) Then
           
        Select Case AlcorMPMessage
            
            Case WM_FT_RW_UNKNOW_FAIL
                TestResult = "Bin2"
                MPTester.TestResultLab = "1St: UnKnow Fail"
                
            Case WM_FT_FLASH_NUM_FAIL
                TestResult = "Bin2"
                MPTester.TestResultLab = "1St: Flash Number Fail"
            
            Case WM_FT_CHECK_HW_CODE_FAIL
                 TestResult = "Bin5"
                 MPTester.TestResultLab = "1St: HW-ID Fail"
            
'            Case WM_FT_TESTUNITREADY_FAIL
'                 TestResult = "Bin2"
'                 MPTester.TestResultLab = "1St: TestUnitReady Fail"
            
            Case WM_FT_RW_SPEED_FAIL
                 TestResult = "Bin3"
                 MPTester.TestResultLab = "1St: SPEED Error "
            
            Case WM_FT_RW_RW_FAIL
                 TestResult = "Bin3"
                 MPTester.TestResultLab = "1St: RW FAIL "
            
            Case WM_FT_CHECK_CERBGPO_FAIL
            
                TestResult = "Bin3"
                MPTester.TestResultLab = "1St: GPO/RB FAIL "
            
            Case WM_FT_CHECK_WRITE_PROTECT_FAIL
                 TestResult = "Bin3"
                 MPTester.TestResultLab = "1St: W/P FAIL "
                 
            Case WM_FT_NO_CARD_FAIL
                  TestResult = "Bin4"
                  MPTester.TestResultLab = "1St: NoCard FAIL "
            
            Case WM_FT_RW_ROM_FAIL
                TestResult = "Bin4"
                MPTester.TestResultLab = "1St: ROM FAIL "
            
            Case WM_FT_RW_RAM_FAIL
                TestResult = "Bin5"
                MPTester.TestResultLab = "1St: RAM FAIL "
            
            Case WM_FT_PHYREAD_FAIL
                TestResult = "Bin5"
                MPTester.TestResultLab = "1St: PHY Read FAIL"
                
            Case WM_FT_NOFREEBLOCK_FAIL
                TestResult = "Bin5"
                MPTester.TestResultLab = "1St: NOFREEBLOCK FAIL"
            
            Case WM_FT_LODECODE_FAIL
                TestResult = "Bin5"
                MPTester.TestResultLab = "1St: LODECODE FAIL"
            
            Case WM_FT_ECC_FAIL
                TestResult = "Bin5"
                MPTester.TestResultLab = "1St: ECC FAIL"
            
'            Case WM_FT_RELOADCODE_FAIL
'                TestResult = "Bin5"
'                MPTester.TestResultLab = "1St: RELOADCODE FAIL"
            
            Case WM_FT_MOVE_DATA_FAIL
                TestResult = "Bin5"
                MPTester.TestResultLab = "1St: MOVE DATA FAIL"
                
            Case WM_FT_RW_RW_PASS
                
                If CheckLED_Flag Then
                    For LedCount = 1 To 20
                        Call MsecDelay(0.1)
                        cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
                        
                        If LightOn = &HEF Or LightOn = &HCF Or LightOn = 223 Then
                            Exit For
                        End If
                    Next LedCount
    
                    MPTester.Print "light="; LightOn
        
                    If (LightOn = &HEF Or LightOn = &HCF Or LightOn = 223) Then
                        MPTester.TestResultLab = "HV: PASS "
                        TestResult = "PASS"
                    Else
                        TestResult = "Bin3"
                    End If
                    
                Else
                    MPTester.TestResultLab = "1St: PASS "
                    TestResult = "PASS"
                End If
                
            Case Else
                 TestResult = "Bin2"
                 MPTester.TestResultLab = "1St: Undefine Fail"
        
        End Select
    
        
        HV_Result = TestResult
        TestResult = ""
        
        SetSiteStatus (HVDone)
        WaitAnotherSiteDone (HVDone)
        
        Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
        cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'Power OFF UNLoad Device
        
        WaitDevOFF ("vid_058f")
        Call MsecDelay(0.3)
        
        SetSiteStatus (RunLV)
        WaitAnotherSiteDone (RunLV)
        
        If Dir("D:\LABPC.PC") = "LABPC.PC" Then
            Call PowerSet2(1, "4.8", "0.15", 1, "4.8", "0.15", 1)
        Else
            Call PowerSet2(1, CurDevicePar.SetLV1, CurDevicePar.SetLI1, 1, CurDevicePar.SetLV2, CurDevicePar.SetLI2, 1)
        End If
        
        EQC_LV = True
        
        If Dual_Flag Then
            cardresult = DO_WritePort(card, Channel_P1A, &HFA)
        Else
            cardresult = DO_WritePort(card, Channel_P1A, &HFB)
        End If
        
        
        WaitDevOn ("vid_058f")
        Call MsecDelay(0.2)
        
        GoTo RW_Test_Label
        
        
        
    ElseIf (EQC_HV = True) And (EQC_LV = True) Then
            
        Select Case AlcorMPMessage
    
        Case WM_FT_RW_UNKNOW_FAIL
            TestResult = "Bin2"
            MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "2nd: UnKnow Fail"
            
        Case WM_FT_FLASH_NUM_FAIL
            TestResult = "Bin2"
            MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "2nd: Flash Number Fail"
            
        Case WM_FT_CHECK_HW_CODE_FAIL
             TestResult = "Bin5"
             MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "2nd: HW-ID Fail"
            
'        Case WM_FT_TESTUNITREADY_FAIL
'             TestResult = "Bin2"
'             MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "2nd: TestUnitReady Fail"
             
        Case WM_FT_RW_SPEED_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "2nd: SPEED Error "
             
        Case WM_FT_RW_RW_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "2nd: RW FAIL "
             
        Case WM_FT_CHECK_CERBGPO_FAIL
    
            TestResult = "Bin3"
            MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "2nd: GPO/RB FAIL "
            
        Case WM_FT_CHECK_WRITE_PROTECT_FAIL
            TestResult = "Bin3"
            MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "2nd: W/P FAIL "
             
        Case WM_FT_NO_CARD_FAIL
            TestResult = "Bin4"
            MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "2nd: NoCard FAIL "
            
        Case WM_FT_RW_ROM_FAIL
            TestResult = "Bin4"
            MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "2nd: ROM FAIL "
            
        Case WM_FT_RW_RAM_FAIL, WM_FT_PHYREAD_FAIL, WM_FT_ECC_FAIL, WM_FT_NOFREEBLOCK_FAIL, WM_FT_LODECODE_FAIL, WM_FT_MOVE_DATA_FAIL
            TestResult = "Bin4"
            MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "2nd: RAM FAIL "
            
        Case WM_FT_RW_RW_PASS
            
            If CheckLED_Flag Then
                For LedCount = 1 To 20
                    Call MsecDelay(0.1)
                    cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
                    
                    If LightOn = &HEF Or LightOn = &HCF Or LightOn = 223 Then
                        Exit For
                    End If
                Next LedCount
    
                MPTester.Print "light="; LightOn
    
                If (LightOn = &HEF Or LightOn = &HCF Or LightOn = 223) Then
                    MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "2nd: PASS"
                    TestResult = "PASS"
                Else
                    TestResult = "Bin3"
                End If
            
            Else
                MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "2nd: PASS"
                TestResult = "PASS"
            End If
        
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "2nd: Undefine Fail"
                   
        End Select
          
        LV_Result = TestResult
        TestResult = ""
        
        If (HV_Result = "Bin2") And (LV_Result = "Bin2") Then
            TestResult = "Bin2"
            ContFail = ContFail + 1
        ElseIf (HV_Result <> "PASS") And (LV_Result = "PASS") Then
            TestResult = "Bin3"
            ContFail = ContFail + 1
        ElseIf (HV_Result = "PASS") And (LV_Result <> "PASS") Then
            TestResult = "Bin4"
            ContFail = ContFail + 1
        ElseIf (HV_Result <> "PASS") And (LV_Result <> "PASS") Then
            TestResult = "Bin5"
            ContFail = ContFail + 1
        ElseIf (HV_Result = "PASS") And (LV_Result = "PASS") Then
            TestResult = "PASS"
            ContFail = 0
        Else
            TestResult = "Bin2"
            ContFail = ContFail + 1
        End If
        
        SetSiteStatus (LVDone)
        WaitAnotherSiteDone (LVDone)
      
    End If
    
TestEnd:
    
    Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
    cardresult = DO_WritePort(card, Channel_P1A, &HFF)
    SetSiteStatus (SiteUnknow)
    WaitDevOFF ("058f")
                       
End Sub



Public Sub AU6919ST4TestSub()

Dim OldTimer
Dim PassTime
Dim rt2
Dim LightOn
Dim mMsg As MSG
Dim LedCount As Byte
Dim tmpName As String
Dim ICName As String, MODELName As String
Dim TimerCounter As Integer
Dim TmpString As String
Dim MP_Retry As Byte
Dim FirstEnum As Boolean


    'add unload driver function
    If PCI7248InitFinish = 0 Then
       PCI7248Exist
    End If
    
    FirstEnum = False
    MP_Retry = 0
    tmpName = Left(ChipName, 9)
    MPTester.TestResultLab = ""
    NewChipFlag = 0
    ChDir (App.Path & CurDevicePar.DeviceFolder)
        
    If OldChipName <> ChipName Then
        '===============================================================
        ' Fail location initial
        '===============================================================
        If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
            FileCopy App.Path & CurDevicePar.DeviceFolder & "\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
            Call MsecDelay(5)
        End If
    
        If (OldVer_Flag) Then
            FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\ROM.Hex", App.Path & CurDevicePar.DeviceFolder & "\ROM.Hex"
            FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\RAM.Bin", App.Path & CurDevicePar.DeviceFolder & "\RAM.Bin"
            FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\AlcorMP.ini", App.Path & CurDevicePar.DeviceFolder & "\AlcorMP.ini"
            FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\PE.bin", App.Path & CurDevicePar.DeviceFolder & "\PE.bin"
            FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\FT.ini", App.Path & CurDevicePar.DeviceFolder & "\FT.ini"
        Else
            FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\ROM.Hex", App.Path & CurDevicePar.DeviceFolder & "\ROM.Hex"
            FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\RAM.Bin", App.Path & CurDevicePar.DeviceFolder & "\RAM.Bin"
            FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\AlcorMP.ini", App.Path & CurDevicePar.DeviceFolder & "\AlcorMP.ini"
            FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\PE.bin", App.Path & CurDevicePar.DeviceFolder & "\PE.bin"
            FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\FT.ini", App.Path & CurDevicePar.DeviceFolder & "\FT.ini"
            If PLFlag = True Then
                FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\PL_FT.ini", App.Path & CurDevicePar.DeviceFolder & "\FT.ini"
            ElseIf KLFlag = True Then
                FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\KL_FT.ini", App.Path & CurDevicePar.DeviceFolder & "\FT.ini"
            Else
                FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\FT.ini", App.Path & CurDevicePar.DeviceFolder & "\FT.ini"
            End If
        End If
        
        NewChipFlag = 1 ' force MP
    End If

    OldChipName = ChipName
    SetSiteStatus (SiteReady)
    
    '==============================================================
    ' when begin RW Test, must clear MP program
    '===============================================================
    Call CloseMP_AU69XX
    MPTester.Print "ContFail="; ContFail
    MPTester.Print "MPContFail="; MPContFail
 
    '====================================
    '  Fix Card
    '====================================
    If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag) Then
        If MPTester.NoMP.Value = 1 Then
            If (NewChipFlag = 0) And (MPTester.Check1.Value = 0) Then  ' force condition
                GoTo RW_Test_Label
            End If
        End If
        
        If MPTester.ResetMPFailCounter.Value = 1 Then
            ContFail = 0
        End If
        
        '===============================================================
        ' when begin MP, must close RW program
        '===============================================================
        MPFlag = 1
        Call CloseFTtool_AU69XX
        
        'power on
        cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)   ' close power to disable chip
        Call MsecDelay(0.5)  ' power for load MPDriver
        MPTester.Print "wait for MP Ready"
        Call LoadMP_AU69XX
    
        OldTimer = Timer
        AlcorMPMessage = 0
        Do
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
            PassTime = Timer - OldTimer
        
        Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 30 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
    
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3
        '====================================================
        If PassTime > 30 Then
        
            Call CloseMP_AU69XX
            MPTester.TestResultLab = "Bin3:MP Ready Fail"
            TestResult = "Bin3"
            MPTester.Print "MP Ready Fail"
            SetSiteStatus (SiteUnknow)
            GoTo TestEnd
        End If
            
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_MP_START Then
            
            SetSiteStatus (RunMP)
            WaitAnotherSiteDone (RunMP)
            
            If Dir("D:\LABPC.PC") = "LABPC.PC" Then
                Call PowerSet2(1, "5.0", "0.5", 1, "5.0", "0.5", 1)
            Else
                Call PowerSet2(1, CurDevicePar.SetStdV1, "0.5", 1, CurDevicePar.SetStdV2, "0.5", 1)
            End If
            
            If Dual_Flag Then
                cardresult = DO_WritePort(card, Channel_P1A, &HFD)
            Else
                cardresult = DO_WritePort(card, Channel_P1A, &HFB)
            End If
                             
            Do
                DoEvents
                Call MsecDelay(0.1)
                TimerCounter = TimerCounter + 1
                TmpString = GetDeviceName("vid")
            Loop While (TmpString = "") And (TimerCounter < 100)
                 
            Call MsecDelay(0.3)
                
            If TmpString = "" Then   ' can not find device after 15 s
                TestResult = "Bin2"
                MPTester.TestResultLab = "Bin2:MP UNKNOW Fail when enter MP"
                SetSiteStatus (SiteUnknow)
                GoTo TestEnd
            End If
                 
            Call MsecDelay(2.5)
                   
            MPTester.Print " MP Begin....."
                 
            Call StartMP_AU69XX
            OldTimer = Timer
            AlcorMPMessage = 0
            ReMP_Flag = 0
            
            Do
                If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                    AlcorMPMessage = mMsg.message
                    TranslateMessage mMsg
                    DispatchMessage mMsg
                    
                    If (AlcorMPMessage = WM_FT_MP_FAIL) And (MP_Retry < 3) Then
                        AlcorMPMessage = 1
                        cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'close power
                        Call MsecDelay(0.3)
                        If Dual_Flag Then
                            cardresult = DO_WritePort(card, Channel_P1A, &HFD)
                        Else
                            cardresult = DO_WritePort(card, Channel_P1A, &HFB)
                        End If
                        Call MsecDelay(2.2)
                        Call RefreshMP_AU69XX
                        Call MsecDelay(0.5)
                        Call StartMP_AU69XX
                        MP_Retry = MP_Retry + 1
                    End If
                End If
                
                PassTime = Timer - OldTimer
            Loop Until AlcorMPMessage = WM_FT_MP_PASS _
                    Or AlcorMPMessage = WM_FT_MP_FAIL _
                    Or AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL _
                    Or PassTime > 65 _
                    Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                    
            MPTester.Print "MP work time="; PassTime
            MPTester.MPText.Text = Hex(AlcorMPMessage)
            
            
            '=========================================
            '    Close MP program
            '=========================================
            Call CloseMP_AU69XX
            KillProcess ("AlcorMP.exe")
            Call UnloadDriver
            
            
            '===============================================
            '  Handle MP work time out error
            '===============================================
            
            'time out fail
            If PassTime > 65 Then
                Call CloseMP_AU69XX
                MPContFail = MPContFail + 1
                TestResult = "Bin3"
                MPTester.TestResultLab = "Bin3:MP Time out Fail"
                MPTester.Print "MP Time out Fail"
                SetSiteStatus (SiteUnknow)
                GoTo TestEnd
            End If
               
            'MP fail
            If AlcorMPMessage = WM_FT_MP_FAIL Then
                Call CloseMP_AU69XX
                MPContFail = MPContFail + 1
                TestResult = "Bin3"
                MPTester.TestResultLab = "Bin3:MP Function Fail"
                MPTester.Print "MP Function Fail"
                SetSiteStatus (SiteUnknow)
                GoTo TestEnd
            End If
        
            'unknow fail
            If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                Call CloseMP_AU69XX
                MPContFail = MPContFail + 1
                TestResult = "Bin2"
                MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                MPTester.Print "MP UNKNOW Fail"
                SetSiteStatus (SiteUnknow)
                GoTo TestEnd
            End If
               
            ' mp pass
            If AlcorMPMessage = WM_FT_MP_PASS Then
                MPTester.TestResultLab = "MP PASS"
                MPContFail = 0
                MPTester.Print "MP PASS"
                SetSiteStatus (MPDone)
            End If
        End If
    End If


    '=========================================
    '    POWER on
    '=========================================
    
RW_Test_Label:
    
    winHwnd = FindWindow(vbNullString, CurDevicePar.FT_ToolTitle)
    
    If winHwnd = 0 Then
        Call LoadFTtool_AU69XX
    
        MPTester.Print "wait for RW Tester Ready"
        OldTimer = Timer
        AlcorMPMessage = 0
        Do
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
            
            PassTime = Timer - OldTimer
        Loop Until AlcorMPMessage = WM_FT_RW_READY Or PassTime > 5 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
        
        MPTester.Print "RW Ready Time="; PassTime
        
        If PassTime > 5 Then
            CloseFTtool_AU69XX
            TestResult = "Bin3"
            MPTester.TestResultLab = "Bin3:RW Ready Fail"
            SetSiteStatus (SiteUnknow)
            GoTo TestEnd
        End If
    End If
    
    If MPFlag = 1 Then
        WaitAnotherSiteDone (MPDone)
        Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
        cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'Power OFF UNLoad Device
        WaitDevOFF ("vid_058f")
        Call MsecDelay(0.3)
        Call PowerSet2(1, CurDevicePar.SetStdV1, CurDevicePar.SetStdI1, 1, CurDevicePar.SetStdV2, CurDevicePar.SetStdI2, 1)
        
        If Dual_Flag Then
            cardresult = DO_WritePort(card, Channel_P1A, &HFE)
            Call MsecDelay(0.04)
            cardresult = DO_WritePort(card, Channel_P1A, &HFA)
        Else
            cardresult = DO_WritePort(card, Channel_P1A, &HFB)
        End If
        
        MPTester.Print "RW Tester begin test........"
        
        FirstEnum = WaitDevOn("vid_058f")
        Call MsecDelay(12#)
        ST4FirstMP = True
        MPFlag = 0
    Else
        Call PowerSet2(1, CurDevicePar.SetStdV1, CurDevicePar.SetStdI1, 1, CurDevicePar.SetStdV2, CurDevicePar.SetStdI2, 1)
        
        If Dual_Flag Then
            cardresult = DO_WritePort(card, Channel_P1A, &HFE)
            Call MsecDelay(0.04)
            cardresult = DO_WritePort(card, Channel_P1A, &HFA)
        Else
            cardresult = DO_WritePort(card, Channel_P1A, &HFB)
        End If
        
        SetSiteStatus (RunHV)
        MPTester.Print "RW Tester begin test........"
                    
        FirstEnum = WaitDevOn("vid_058f")
        Call MsecDelay(0.3)
        ST4FirstMP = False
        
    End If
        
    If Not FirstEnum Then
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2:Unknow Fail"
        SetSiteStatus (SiteUnknow)
        GoTo TestEnd
    Else
        SetSiteStatus (RunLV)
        WaitAnotherSiteDone (RunLV)
        Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
        cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'Power OFF UNLoad Device
        WaitDevOFF ("vid_058f")
        Call MsecDelay(0.2)
        Call PowerSet2(1, CurDevicePar.SetStdV1, CurDevicePar.SetStdI1, 1, CurDevicePar.SetStdV2, CurDevicePar.SetStdI2, 1)
        
        If Dual_Flag Then
            cardresult = DO_WritePort(card, Channel_P1A, &HFE)
            Call MsecDelay(0.04)
            cardresult = DO_WritePort(card, Channel_P1A, &HFA)
        Else
            cardresult = DO_WritePort(card, Channel_P1A, &HFB)
        End If
        
        
        If (CurDevicePar.ShortName = "AU6992") Then
            WaitDevOn ("vid_058f")
            Call MsecDelay(0.3)
            Call StartFTtest_AU69XX
        Else
            Call StartFTtest_WaitDevReady_AU69XX(4#, False)
        End If
       
        
        'SetSiteStatus (LVDone)
        'WaitAnotherSiteDone (LVDone)
    End If
    
    AlcorMPMessage = 0
    SetSiteStatus (LVDone)
    'MPTester.Print "RW Tester begin test........"
    OldTimer = Timer
    
'    If (EQC_HV = True) And (EQC_LV = True) Then
'        Call StartFTtest_WaitDevReady_AU69XX(0.5)
'    End If
    
    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
        
        PassTime = Timer - OldTimer
       
    Loop Until AlcorMPMessage = WM_FT_RW_SPEED_FAIL _
            Or AlcorMPMessage = WM_FT_RW_RW_FAIL _
            Or AlcorMPMessage = WM_FT_RW_ROM_FAIL _
            Or AlcorMPMessage = WM_FT_RW_RAM_FAIL _
            Or AlcorMPMessage = WM_FT_RW_RW_PASS _
            Or AlcorMPMessage = WM_FT_RW_UNKNOW_FAIL _
            Or AlcorMPMessage = WM_FT_FLASH_NUM_FAIL _
            Or AlcorMPMessage = WM_FT_CHECK_CERBGPO_FAIL _
            Or AlcorMPMessage = WM_FT_CHECK_HW_CODE_FAIL _
            Or AlcorMPMessage = WM_FT_PHYREAD_FAIL _
            Or AlcorMPMessage = WM_FT_ECC_FAIL _
            Or AlcorMPMessage = WM_FT_NOFREEBLOCK_FAIL _
            Or AlcorMPMessage = WM_FT_LODECODE_FAIL _
            Or AlcorMPMessage = WM_FT_CHECK_WRITE_PROTECT_FAIL _
            Or AlcorMPMessage = WM_FT_NO_CARD_FAIL _
            Or AlcorMPMessage = WM_FT_LOCK_VOLUME_FAIL _
            Or AlcorMPMessage = WM_FT_UNLOCK_VOLUME_FAIL _
            Or AlcorMPMessage = WM_FT_TURNOFF_TUR_FAIL _
            Or AlcorMPMessage = WM_DEV_GET_HANDLE_FAIL _
            Or AlcorMPMessage = WM_DEV_GET_DIS_TUR_HANDLE_FAIL _
            Or PassTime > 10 _
            Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                    
    MPTester.Print "RW work Time="; PassTime
    MPTester.MPText.Text = Hex(AlcorMPMessage)
            
    '===========================================================
    '  RW Time Out Fail
    '===========================================================
    
    If (PassTime > 10) Or ((FailCloseAP) And (AlcorMPMessage <> WM_FT_RW_RW_PASS)) Then
        Close_FT_AP ("UFD Test")
        
        If (PassTime > 8) Then
            TestResult = "Bin3"
            MPTester.TestResultLab = "Bin3:RW Time Out Fail"
            cardresult = DO_WritePort(card, Channel_P1A, &HFF)  ' power off
            SetSiteStatus (SiteUnknow)
            WaitAnotherSiteDone (HVDone)
            GoTo TestEnd
        End If
    End If
                        
    Select Case AlcorMPMessage
        Case WM_FT_RW_UNKNOW_FAIL
            TestResult = "Bin2"
            MPTester.TestResultLab = "Bin2:UnKnow Fail"
            ContFail = ContFail + 1
            
        Case WM_FT_FLASH_NUM_FAIL
            TestResult = "Bin2"
            MPTester.TestResultLab = "Bin2:Flash Number Fail"
            ContFail = ContFail + 1
        
        Case WM_FT_CHECK_HW_CODE_FAIL
            TestResult = "Bin5"
            MPTester.TestResultLab = "Bin5:HW-ID Fail"
            ContFail = ContFail + 1
        
'        Case WM_FT_TESTUNITREADY_FAIL
'            TestResult = "Bin2"
'            MPTester.TestResultLab = "Bin2:TestUnitReady Fail"
'            ContFail = ContFail + 1
        
        Case WM_FT_RW_SPEED_FAIL
            TestResult = "Bin3"
            MPTester.TestResultLab = "Bin3:SPEED Error "
            ContFail = ContFail + 1
             
        Case WM_FT_RW_RW_FAIL
            TestResult = "Bin3"
            MPTester.TestResultLab = "Bin3:RW FAIL "
            ContFail = ContFail + 1
        
        Case WM_FT_CHECK_CERBGPO_FAIL
            TestResult = "Bin3"
            MPTester.TestResultLab = "Bin3:GPO/RB FAIL "
            ContFail = ContFail + 1
        
        Case WM_FT_CHECK_WRITE_PROTECT_FAIL
            TestResult = "Bin3"
            MPTester.TestResultLab = "Bin3:W/P FAIL "
            ContFail = ContFail + 1
             
        Case WM_FT_NO_CARD_FAIL
            TestResult = "Bin4"
            MPTester.TestResultLab = "Bin4:NoCard FAIL "
            ContFail = ContFail + 1
        
        Case WM_FT_RW_ROM_FAIL
            TestResult = "Bin4"
            MPTester.TestResultLab = "Bin4:ROM FAIL "
            ContFail = ContFail + 1
              
        Case WM_FT_PHYREAD_FAIL
            TestResult = "Bin4"
            MPTester.TestResultLab = "Bin4:PHY Read FAIL "
            ContFail = ContFail + 1
              
        Case WM_FT_RW_RAM_FAIL
            TestResult = "Bin4"
            MPTester.TestResultLab = "Bin4:RAM FAIL "
            ContFail = ContFail + 1
               
        Case WM_FT_NOFREEBLOCK_FAIL
            TestResult = "Bin4"
            MPTester.TestResultLab = "Bin4:FreeBlock FAIL "
            ContFail = ContFail + 1
        
        Case WM_FT_LODECODE_FAIL
            TestResult = "Bin4"
            MPTester.TestResultLab = "Bin4:LoadCode FAIL "
            ContFail = ContFail + 1
        
'        Case WM_FT_RELOADCODE_FAIL
'            TestResult = "Bin4"
'            MPTester.TestResultLab = "Bin4:ReLoadCode FAIL "
'            ContFail = ContFail + 1
        
        Case WM_FT_ECC_FAIL
            TestResult = "Bin5"
            MPTester.TestResultLab = "Bin5:ECC FAIL "
            ContFail = ContFail + 1
            
        Case WM_FT_MOVE_DATA_FAIL
            TestResult = "Bin5"
            MPTester.TestResultLab = "Bin5:MOVE DATA FAIL "
            ContFail = ContFail + 1
                     
        Case WM_DEV_GET_HANDLE_FAIL
            TestResult = "Bin2"
            MPTester.TestResultLab = "Bin2:Get Handle FAIL "
            ContFail = ContFail + 1
            
        Case WM_DEV_GET_DIS_TUR_HANDLE_FAIL
            TestResult = "Bin2"
            MPTester.TestResultLab = "Bin2:TURN ON TUR FAIL "
            ContFail = ContFail + 1
        
        Case WM_FT_LOCK_VOLUME_FAIL
           TestResult = "Bin2"
           MPTester.TestResultLab = "Bin2:LOCK VOLUME FAIL "
           ContFail = ContFail + 1
        
        Case WM_FT_UNLOCK_VOLUME_FAIL
           TestResult = "Bin2"
           MPTester.TestResultLab = "Bin2:UNLOCK VOLUME FAIL "
           ContFail = ContFail + 1
           
        Case WM_FT_TURNOFF_TUR_FAIL
           TestResult = "Bin2"
           MPTester.TestResultLab = "Bin2:TURN OFF TUR FAIL "
           ContFail = ContFail + 1
                    
                    
        Case WM_FT_RW_RW_PASS
            If CheckLED_Flag Then
                For LedCount = 1 To 20
                    Call MsecDelay(0.1)
                    cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
                    If LightOn = &HEF Or LightOn = &HCF Or LightOn = 223 Then
                        Exit For
                    End If
                Next LedCount
                
                MPTester.Print "light="; LightOn
                
                If LightOn = &HEF Or LightOn = &HCF Or LightOn = 223 Then
                    MPTester.TestResultLab = "PASS "
                    TestResult = "PASS"
                    ContFail = 0
                Else
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:LED FAIL "
                    ContFail = ContFail + 1
                End If
            Else
                MPTester.TestResultLab = "PASS "
                TestResult = "PASS"
                ContFail = 0
            End If
            
        Case Else
            TestResult = "Bin2"
            MPTester.TestResultLab = "Bin2:Undefine Fail"
            ContFail = ContFail + 1
               
    End Select
    
    SetSiteStatus (HVDone)
    WaitAnotherSiteDone (HVDone)
    
TestEnd:
    
    cardresult = DO_WritePort(card, Channel_P1A, &HFF)
    Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
    SetSiteStatus (SiteUnknow)
    WaitDevOFF ("058f")
    
End Sub

Public Sub AU6913ST1TestSub()

'2012/6/9: purpose to solve Transcent RMA,
'modify power-on sequence(1. V33 => 2. Ena) to kill this power sequence unknow sample.
'because LDO power-on sequence were random.

Dim OldTimer
Dim PassTime
Dim rt2
Dim LightOn
Dim mMsg As MSG
Dim LedCount As Byte
Dim tmpName As String
Dim ICName As String, MODELName As String
Dim TimerCounter As Integer
Dim TmpString As String
Dim MP_Retry As Byte
Dim FirstEnum As Boolean


If PCI7248InitFinish = 0 Then
   PCI7248Exist
End If

FirstEnum = False
MP_Retry = 0
tmpName = Left(ChipName, 9)
MPTester.TestResultLab = ""
NewChipFlag = 0
ChDir (App.Path & CurDevicePar.DeviceFolder)
    
If OldChipName <> ChipName Then
    '===============================================================
    ' Fail location initial
    '===============================================================
    If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
        FileCopy App.Path & CurDevicePar.DeviceFolder & "\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
        Call MsecDelay(5)
    End If

    If (OldVer_Flag) Then
        FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\ROM.Hex", App.Path & CurDevicePar.DeviceFolder & "\ROM.Hex"
        FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\RAM.Bin", App.Path & CurDevicePar.DeviceFolder & "\RAM.Bin"
        FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\AlcorMP.ini", App.Path & CurDevicePar.DeviceFolder & "\AlcorMP.ini"
        FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\PE.bin", App.Path & CurDevicePar.DeviceFolder & "\PE.bin"
        FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\FT.ini", App.Path & CurDevicePar.DeviceFolder & "\FT.ini"
    Else
        FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\ROM.Hex", App.Path & CurDevicePar.DeviceFolder & "\ROM.Hex"
        FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\RAM.Bin", App.Path & CurDevicePar.DeviceFolder & "\RAM.Bin"
        FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\AlcorMP.ini", App.Path & CurDevicePar.DeviceFolder & "\AlcorMP.ini"
        FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\PE.bin", App.Path & CurDevicePar.DeviceFolder & "\PE.bin"
        
        If PLFlag = True Then
            FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\PL_FT.ini", App.Path & CurDevicePar.DeviceFolder & "\FT.ini"
        ElseIf KLFlag = True Then
            FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\KL_FT.ini", App.Path & CurDevicePar.DeviceFolder & "\FT.ini"
        Else
            FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\FT.ini", App.Path & CurDevicePar.DeviceFolder & "\FT.ini"
        End If
                
    End If
    
    NewChipFlag = 1 ' force MP
End If

OldChipName = ChipName
SetSiteStatus (SiteReady)
    
'==============================================================
' when begin RW Test, must clear MP program
'===============================================================
Call CloseMP_AU69XX
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail
 
'====================================
'  Fix Card
'====================================
If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag) Then
    If MPTester.NoMP.Value = 1 Then
        If (NewChipFlag = 0) And (MPTester.Check1.Value = 0) Then  ' force condition
            GoTo RW_Test_Label
        End If
    End If
    
    If MPTester.ResetMPFailCounter.Value = 1 Then
        ContFail = 0
    End If
    
    '===============================================================
    ' when begin MP, must close RW program
    '===============================================================
    MPFlag = 1
    Call CloseFTtool_AU69XX
    
    'power on
    cardresult = DO_WritePort(card, Channel_P1A, &HFF)
    Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)   ' close power to disable chip
    Call MsecDelay(0.5)  ' power for load MPDriver
    MPTester.Print "wait for MP Ready"
    Call LoadMP_AU69XX

    OldTimer = Timer
    AlcorMPMessage = 0
    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
        PassTime = Timer - OldTimer
    
    Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 30 _
            Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY

    MPTester.Print "Ready Time="; PassTime
    
    '====================================================
    '  handle MP load time out, the FAIL will be Bin3
    '====================================================
    If PassTime > 30 Then
    
        Call CloseMP_AU69XX
        MPTester.TestResultLab = "Bin3:MP Ready Fail"
        TestResult = "Bin3"
        MPTester.Print "MP Ready Fail"
        SetSiteStatus (SiteUnknow)
        GoTo TestEnd
    End If
        
    '====================================================
    '  MP begin
    '====================================================
    
    If AlcorMPMessage = WM_FT_MP_START Then
        
        SetSiteStatus (RunMP)
        WaitAnotherSiteDone (RunMP)
        
        If Dual_Flag Then
            cardresult = DO_WritePort(card, Channel_P1A, &HFD)
        Else
            cardresult = DO_WritePort(card, Channel_P1A, &HFB)
        End If
        
        Call PowerSet2(1, CurDevicePar.SetStdV1, "0.5", 1, CurDevicePar.SetStdV2, "0.5", 1)
        
        Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName("vid")
        Loop While (TmpString = "") And (TimerCounter < 100)
             
        Call MsecDelay(0.3)
            
        If TmpString = "" Then   ' can not find device after 15 s
            TestResult = "Bin2"
            MPTester.TestResultLab = "Bin2:MP UNKNOW Fail when enter MP"
            SetSiteStatus (SiteUnknow)
            GoTo TestEnd
        End If
             
        Call MsecDelay(2.5)
               
        MPTester.Print " MP Begin....."
             
        Call StartMP_AU69XX
        OldTimer = Timer
        AlcorMPMessage = 0
        ReMP_Flag = 0
        
        Do
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
                
                If (AlcorMPMessage = WM_FT_MP_FAIL) And (MP_Retry < 3) Then
                    AlcorMPMessage = 1
                    cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'close power
                    Call MsecDelay(0.3)
                    If Dual_Flag Then
                        cardresult = DO_WritePort(card, Channel_P1A, &HFD)
                    Else
                        cardresult = DO_WritePort(card, Channel_P1A, &HFB)
                    End If
                    Call MsecDelay(2.2)
                    Call RefreshMP_AU69XX
                    Call MsecDelay(0.5)
                    Call StartMP_AU69XX
                    MP_Retry = MP_Retry + 1
                End If
            End If
            
            PassTime = Timer - OldTimer
        Loop Until AlcorMPMessage = WM_FT_MP_PASS _
                Or AlcorMPMessage = WM_FT_MP_FAIL _
                Or AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL _
                Or PassTime > 65 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                
        MPTester.Print "MP work time="; PassTime
        MPTester.MPText.Text = Hex(AlcorMPMessage)
        
        
        '=========================================
        '    Close MP program
        '=========================================
        Call CloseMP_AU69XX
        KillProcess ("AlcorMP.exe")
        Call UnloadDriver
        
        
        '===============================================
        '  Handle MP work time out error
        '===============================================
        
        'time out fail
        If PassTime > 65 Then
            Call CloseMP_AU69XX
            MPContFail = MPContFail + 1
            TestResult = "Bin3"
            MPTester.TestResultLab = "Bin3:MP Time out Fail"
            MPTester.Print "MP Time out Fail"
            SetSiteStatus (SiteUnknow)
            GoTo TestEnd
        End If
           
        'MP fail
        If AlcorMPMessage = WM_FT_MP_FAIL Then
            Call CloseMP_AU69XX
            MPContFail = MPContFail + 1
            TestResult = "Bin3"
            MPTester.TestResultLab = "Bin3:MP Function Fail"
            MPTester.Print "MP Function Fail"
            SetSiteStatus (SiteUnknow)
            GoTo TestEnd
        End If
    
        'unknow fail
        If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
            Call CloseMP_AU69XX
            MPContFail = MPContFail + 1
            TestResult = "Bin2"
            MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
            MPTester.Print "MP UNKNOW Fail"
            SetSiteStatus (SiteUnknow)
            GoTo TestEnd
        End If
           
        ' mp pass
        If AlcorMPMessage = WM_FT_MP_PASS Then
            MPTester.TestResultLab = "MP PASS"
            MPContFail = 0
            MPTester.Print "MP PASS"
            SetSiteStatus (MPDone)
        End If
    End If
End If


'=========================================
'    POWER on
'=========================================
    
RW_Test_Label:
    
winHwnd = FindWindow(vbNullString, CurDevicePar.FT_ToolTitle)

If winHwnd = 0 Then
    Call LoadFTtool_AU69XX

    MPTester.Print "wait for RW Tester Ready"
    OldTimer = Timer
    AlcorMPMessage = 0
    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
        
        PassTime = Timer - OldTimer
    Loop Until AlcorMPMessage = WM_FT_RW_READY Or PassTime > 5 _
            Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
    
    MPTester.Print "RW Ready Time="; PassTime
    
    If PassTime > 5 Then
        CloseFTtool_AU69XX
        TestResult = "Bin3"
        MPTester.TestResultLab = "Bin3:RW Ready Fail"
        SetSiteStatus (SiteUnknow)
        GoTo TestEnd
    End If
End If

If MPFlag = 1 Then
    WaitAnotherSiteDone (MPDone)
    cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'Power OFF UNLoad Device
    Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
    WaitDevOFF ("vid_058f")
    Call MsecDelay(0.1)
    Call PowerSet2(1, CurDevicePar.SetStdV1, CurDevicePar.SetStdI1, 1, CurDevicePar.SetStdV2, CurDevicePar.SetStdI2, 1)
    If Dual_Flag Then
        cardresult = DO_WritePort(card, Channel_P1A, &HFE)
        Call MsecDelay(0.04)
        cardresult = DO_WritePort(card, Channel_P1A, &HFA)
    Else
        cardresult = DO_WritePort(card, Channel_P1A, &HFB)
    End If
    
    FirstEnum = WaitDevOn("vid_058f")
    Call MsecDelay(0.2)
    MPFlag = 0
Else
    Call PowerSet2(1, CurDevicePar.SetStdV1, CurDevicePar.SetStdI1, 1, CurDevicePar.SetStdV2, CurDevicePar.SetStdI2, 1)
    If Dual_Flag Then
        cardresult = DO_WritePort(card, Channel_P1A, &HFE)
        Call MsecDelay(0.04)
        cardresult = DO_WritePort(card, Channel_P1A, &HFA)
    Else
        cardresult = DO_WritePort(card, Channel_P1A, &HFB)
    End If
    
    FirstEnum = WaitDevOn("vid_058f")
    Call MsecDelay(0.1)
    
End If
        
If Not FirstEnum Then
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2:Unknow Fail"
    SetSiteStatus (SiteUnknow)
    GoTo TestEnd
Else
    SetSiteStatus (RunLV)
    WaitAnotherSiteDone (RunLV)
    'Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
    cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'Power OFF UNLoad Device
    WaitDevOFF ("vid_058f")
    Call MsecDelay(0.6)
    'Call PowerSet2(1, CurDevicePar.SetStdV1, CurDevicePar.SetStdI1, 1, CurDevicePar.SetStdV2, CurDevicePar.SetStdI2, 1)
    'Call MsecDelay(0.2)
    If Dual_Flag Then
        cardresult = DO_WritePort(card, Channel_P1A, &HFE)
        Call MsecDelay(0.04)
        cardresult = DO_WritePort(card, Channel_P1A, &HFA)
    Else
        cardresult = DO_WritePort(card, Channel_P1A, &HFB)
    End If
    
    FirstEnum = WaitDevOn("vid_058f")
    Call MsecDelay(0.2)
    
    'SetSiteStatus (LVDone)
    'WaitAnotherSiteDone (LVDone)
End If
    
OldTimer = Timer
AlcorMPMessage = 0
SetSiteStatus (RunHV)
MPTester.Print "RW Tester begin test........"
Call StartFTtest_AU69XX
    
Do
    If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
        AlcorMPMessage = mMsg.message
        TranslateMessage mMsg
        DispatchMessage mMsg
    End If
    
    PassTime = Timer - OldTimer
   
Loop Until AlcorMPMessage = WM_FT_RW_SPEED_FAIL _
        Or AlcorMPMessage = WM_FT_RW_RW_FAIL _
        Or AlcorMPMessage = WM_FT_RW_ROM_FAIL _
        Or AlcorMPMessage = WM_FT_RW_RAM_FAIL _
        Or AlcorMPMessage = WM_FT_RW_RW_PASS _
        Or AlcorMPMessage = WM_FT_RW_UNKNOW_FAIL _
        Or AlcorMPMessage = WM_FT_FLASH_NUM_FAIL _
        Or AlcorMPMessage = WM_FT_CHECK_CERBGPO_FAIL _
        Or AlcorMPMessage = WM_FT_CHECK_HW_CODE_FAIL _
        Or AlcorMPMessage = WM_FT_PHYREAD_FAIL _
        Or AlcorMPMessage = WM_FT_ECC_FAIL _
        Or AlcorMPMessage = WM_FT_NOFREEBLOCK_FAIL _
        Or AlcorMPMessage = WM_FT_LODECODE_FAIL _
        Or AlcorMPMessage = WM_FT_CHECK_WRITE_PROTECT_FAIL _
        Or AlcorMPMessage = WM_FT_NO_CARD_FAIL _
        Or AlcorMPMessage = WM_FT_TEST_DQS_FAIL Or AlcorMPMessage = WM_LC_FAIL Or AlcorMPMessage = WM_RC_FAIL Or PassTime > 10 _
        Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                
MPTester.Print "RW work Time="; PassTime
MPTester.MPText.Text = Hex(AlcorMPMessage)
            
'===========================================================
'  RW Time Out Fail
'===========================================================

If (PassTime > 10) Or ((FailCloseAP) And (AlcorMPMessage <> WM_FT_RW_RW_PASS)) Then
    Close_FT_AP ("UFD Test")
    
    If (PassTime > 10) Then
        TestResult = "Bin3"
        MPTester.TestResultLab = "Bin3:RW Time Out Fail"
        cardresult = DO_WritePort(card, Channel_P1A, &HFF)  ' power off
        SetSiteStatus (SiteUnknow)
        WaitAnotherSiteDone (HVDone)
        GoTo TestEnd
    End If
End If
                    
Select Case AlcorMPMessage
    Case WM_FT_RW_UNKNOW_FAIL
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2:UnKnow Fail"
        ContFail = ContFail + 1
        
    Case WM_FT_FLASH_NUM_FAIL
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2:Flash Number Fail"
        ContFail = ContFail + 1
    
    Case WM_FT_CHECK_HW_CODE_FAIL
        TestResult = "Bin5"
        MPTester.TestResultLab = "Bin5:HW-ID Fail"
        ContFail = ContFail + 1
    
'    Case WM_FT_TESTUNITREADY_FAIL
'        TestResult = "Bin2"
'        MPTester.TestResultLab = "Bin2:TestUnitReady Fail"
'        ContFail = ContFail + 1
    
    Case WM_FT_RW_SPEED_FAIL
        TestResult = "Bin3"
        MPTester.TestResultLab = "Bin3:SPEED Error "
        ContFail = ContFail + 1
         
    Case WM_FT_RW_RW_FAIL
        TestResult = "Bin3"
        MPTester.TestResultLab = "Bin3:RW FAIL "
        ContFail = ContFail + 1
    
    Case WM_FT_CHECK_CERBGPO_FAIL
        TestResult = "Bin3"
        MPTester.TestResultLab = "Bin3:GPO/RB FAIL "
        ContFail = ContFail + 1
    
    Case WM_FT_CHECK_WRITE_PROTECT_FAIL
        TestResult = "Bin3"
        MPTester.TestResultLab = "Bin3:W/P FAIL "
        ContFail = ContFail + 1
         
    Case WM_FT_NO_CARD_FAIL
        TestResult = "Bin4"
        MPTester.TestResultLab = "Bin4:NoCard FAIL "
        ContFail = ContFail + 1
    
    Case WM_FT_RW_ROM_FAIL
        TestResult = "Bin4"
        MPTester.TestResultLab = "Bin4:ROM FAIL "
        ContFail = ContFail + 1
          
    Case WM_FT_PHYREAD_FAIL
        TestResult = "Bin4"
        MPTester.TestResultLab = "Bin4:PHY Read FAIL "
        ContFail = ContFail + 1
          
    Case WM_FT_RW_RAM_FAIL
        TestResult = "Bin4"
        MPTester.TestResultLab = "Bin4:RAM FAIL "
        ContFail = ContFail + 1
           
    Case WM_FT_NOFREEBLOCK_FAIL
        TestResult = "Bin4"
        MPTester.TestResultLab = "Bin4:FreeBlock FAIL "
        ContFail = ContFail + 1
    
    Case WM_FT_LODECODE_FAIL
        TestResult = "Bin4"
        MPTester.TestResultLab = "Bin4:LoadCode FAIL "
        ContFail = ContFail + 1
    
'    Case WM_FT_RELOADCODE_FAIL
'        TestResult = "Bin4"
'        MPTester.TestResultLab = "Bin4:ReLoadCode FAIL "
'        ContFail = ContFail + 1
    
    Case WM_FT_ECC_FAIL
        TestResult = "Bin5"
        MPTester.TestResultLab = "Bin5:ECC FAIL "
        ContFail = ContFail + 1
        
    Case WM_FT_MOVE_DATA_FAIL
        TestResult = "Bin5"
        MPTester.TestResultLab = "Bin5:MOVE DATA FAIL "
        ContFail = ContFail + 1
                
    Case WM_FT_RW_RW_PASS
        If CheckLED_Flag Then
            For LedCount = 1 To 20
                Call MsecDelay(0.1)
                cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
                If LightOn = &HEF Or LightOn = &HCF Or LightOn = 223 Then
                    Exit For
                End If
            Next LedCount
            
            MPTester.Print "light="; LightOn
            
            If LightOn = &HEF Or LightOn = &HCF Or LightOn = 223 Then
                MPTester.TestResultLab = "PASS "
                TestResult = "PASS"
                ContFail = 0
            Else
                TestResult = "Bin3"
                MPTester.TestResultLab = "Bin3:LED FAIL "
                ContFail = ContFail + 1
            End If
        Else
            MPTester.TestResultLab = "PASS "
            TestResult = "PASS"
            ContFail = 0
        End If
        
    Case Else
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2:Undefine Fail"
        ContFail = ContFail + 1
           
End Select

SetSiteStatus (HVDone)
WaitAnotherSiteDone (HVDone)
    
TestEnd:
    
cardresult = DO_WritePort(card, Channel_P1A, &HFF)
Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
SetSiteStatus (SiteUnknow)
WaitDevOFF ("058f")
    
End Sub

Public Sub AU69XXLC_ST1TestSub()

'2012/6/9: purpose to solve Transcent RMA,
'modify power-on sequence(1. V33 => 2. Ena) to kill this power sequence unknow sample.
'because LDO power-on sequence were random.

Dim OldTimer
Dim PassTime
Dim rt2
Dim LightOn
Dim mMsg As MSG
Dim LedCount As Byte
Dim tmpName As String
Dim ICName As String, MODELName As String
Dim TimerCounter As Integer
Dim TmpString As String
Dim MP_Retry As Byte


If PCI7248InitFinish = 0 Then
   PCI7248Exist
End If

MP_Retry = 0
tmpName = Left(ChipName, 9)
MPTester.TestResultLab = ""
NewChipFlag = 0
ChDir (App.Path & CurDevicePar.DeviceFolder)
    
If OldChipName <> ChipName Then
    '===============================================================
    ' Fail location initial
    '===============================================================
    If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
        FileCopy App.Path & CurDevicePar.DeviceFolder & "\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
        Call MsecDelay(5)
    End If

    FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\ROM.Hex", App.Path & CurDevicePar.DeviceFolder & "\ROM.Hex"
    FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\RAM.Bin", App.Path & CurDevicePar.DeviceFolder & "\RAM.Bin"
    FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\AlcorMP.ini", App.Path & CurDevicePar.DeviceFolder & "\AlcorMP.ini"
    FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\PE.bin", App.Path & CurDevicePar.DeviceFolder & "\PE.bin"
    FileCopy App.Path & CurDevicePar.DeviceFolder & "\INI\" & tmpName & "\FT.ini", App.Path & CurDevicePar.DeviceFolder & "\FT.ini"
    
    NewChipFlag = 1 ' force MP
End If

OldChipName = ChipName
SetSiteStatus (SiteReady)
    
'==============================================================
' when begin RW Test, must clear MP program
'===============================================================
Call CloseMP_AU69XX
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail
 
'====================================
'  Fix Card
'====================================
If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag) Then
    If MPTester.NoMP.Value = 1 Then
        If (NewChipFlag = 0) And (MPTester.Check1.Value = 0) Then  ' force condition
            GoTo RW_Test_Label
        End If
    End If
    
    If MPTester.ResetMPFailCounter.Value = 1 Then
        ContFail = 0
    End If
    
    '===============================================================
    ' when begin MP, must close RW program
    '===============================================================
    MPFlag = 1
    Call CloseFTtool_AU69XX
    
    'power on
    cardresult = DO_WritePort(card, Channel_P1A, &HFF)
    Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)   ' close power to disable chip
    Call MsecDelay(0.5)  ' power for load MPDriver
    MPTester.Print "wait for MP Ready"
    Call LoadMP_AU69XX

    OldTimer = Timer
    AlcorMPMessage = 0
    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
        PassTime = Timer - OldTimer
    
    Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 30 _
            Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY

    MPTester.Print "Ready Time="; PassTime
    
    '====================================================
    '  handle MP load time out, the FAIL will be Bin3
    '====================================================
    If PassTime > 30 Then
    
        Call CloseMP_AU69XX
        MPTester.TestResultLab = "Bin3:MP Ready Fail"
        TestResult = "Bin3"
        MPTester.Print "MP Ready Fail"
        SetSiteStatus (SiteUnknow)
        GoTo TestEnd
    End If
        
    '====================================================
    '  MP begin
    '====================================================
    
    If AlcorMPMessage = WM_FT_MP_START Then
        
        SetSiteStatus (RunMP)
        WaitAnotherSiteDone (RunMP)
        
        If Dual_Flag Then
            cardresult = DO_WritePort(card, Channel_P1A, &HFD)
        Else
            cardresult = DO_WritePort(card, Channel_P1A, &HFB)
        End If
        
        Call PowerSet2(1, CurDevicePar.SetStdV1, "0.5", 1, "0.0", "0.5", 1)
        Call MsecDelay(0.1)
        Call PowerSet2(0, CurDevicePar.SetStdV1, "0.5", 1, CurDevicePar.SetStdV2, "0.5", 1)
        
        Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName("vid")
        Loop While (TmpString = "") And (TimerCounter < 100)
             
        Call MsecDelay(0.3)
            
        If TmpString = "" Then   ' can not find device after 15 s
            TestResult = "Bin2"
            MPTester.TestResultLab = "Bin2:MP UNKNOW Fail when enter MP"
            SetSiteStatus (SiteUnknow)
            GoTo TestEnd
        End If
             
        Call MsecDelay(2.5)
               
        MPTester.Print " MP Begin....."
             
        Call StartMP_AU69XX
        OldTimer = Timer
        AlcorMPMessage = 0
        ReMP_Flag = 0
        
        Do
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
                
                If (AlcorMPMessage = WM_FT_MP_FAIL) And (MP_Retry < 3) Then
                    AlcorMPMessage = 1
                    cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'close power
                    Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
                    Call MsecDelay(0.3)
                    If Dual_Flag Then
                        cardresult = DO_WritePort(card, Channel_P1A, &HFD)
                    Else
                        cardresult = DO_WritePort(card, Channel_P1A, &HFB)
                    End If
                    Call PowerSet2(1, "0.0", "0.5", 1, CurDevicePar.SetStdV2, "0.5", 1)
                    Call MsecDelay(0.1)
                    Call PowerSet2(0, CurDevicePar.SetStdV1, "0.5", 1, CurDevicePar.SetStdV2, "0.5", 1)
                    Call MsecDelay(2.2)
                    Call RefreshMP_AU69XX
                    Call MsecDelay(0.5)
                    Call StartMP_AU69XX
                    MP_Retry = MP_Retry + 1
                End If
            End If
            
            PassTime = Timer - OldTimer
        Loop Until AlcorMPMessage = WM_FT_MP_PASS _
                Or AlcorMPMessage = WM_FT_MP_FAIL _
                Or AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL _
                Or PassTime > 65 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                
        MPTester.Print "MP work time="; PassTime
        MPTester.MPText.Text = Hex(AlcorMPMessage)
        
        
        '=========================================
        '    Close MP program
        '=========================================
        Call CloseMP_AU69XX
        KillProcess ("AlcorMP.exe")
        Call UnloadDriver
        
        
        '===============================================
        '  Handle MP work time out error
        '===============================================
        
        'time out fail
        If PassTime > 65 Then
            Call CloseMP_AU69XX
            MPContFail = MPContFail + 1
            TestResult = "Bin3"
            MPTester.TestResultLab = "Bin3:MP Time out Fail"
            MPTester.Print "MP Time out Fail"
            SetSiteStatus (SiteUnknow)
            GoTo TestEnd
        End If
           
        'MP fail
        If AlcorMPMessage = WM_FT_MP_FAIL Then
            Call CloseMP_AU69XX
            MPContFail = MPContFail + 1
            TestResult = "Bin3"
            MPTester.TestResultLab = "Bin3:MP Function Fail"
            MPTester.Print "MP Function Fail"
            SetSiteStatus (SiteUnknow)
            GoTo TestEnd
        End If
    
        'unknow fail
        If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
            Call CloseMP_AU69XX
            MPContFail = MPContFail + 1
            TestResult = "Bin2"
            MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
            MPTester.Print "MP UNKNOW Fail"
            SetSiteStatus (SiteUnknow)
            GoTo TestEnd
        End If
           
        ' mp pass
        If AlcorMPMessage = WM_FT_MP_PASS Then
            MPTester.TestResultLab = "MP PASS"
            MPContFail = 0
            MPTester.Print "MP PASS"
            SetSiteStatus (MPDone)
        End If
    End If
End If


'=========================================
'    POWER on
'=========================================
    
RW_Test_Label:
    
winHwnd = FindWindow(vbNullString, CurDevicePar.FT_ToolTitle)

If winHwnd = 0 Then
    Call LoadFTtool_AU69XX

    MPTester.Print "wait for RW Tester Ready"
    OldTimer = Timer
    AlcorMPMessage = 0
    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
        
        PassTime = Timer - OldTimer
    Loop Until AlcorMPMessage = WM_FT_RW_READY Or PassTime > 5 _
            Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
    
    MPTester.Print "RW Ready Time="; PassTime
    
    If PassTime > 5 Then
        CloseFTtool_AU69XX
        TestResult = "Bin3"
        MPTester.TestResultLab = "Bin3:RW Ready Fail"
        SetSiteStatus (SiteUnknow)
        GoTo TestEnd
    End If
End If

If MPFlag = 1 Then
    WaitAnotherSiteDone (MPDone)
    cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'Power OFF UNLoad Device
    Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
    WaitDevOFF ("vid_058f")
    Call MsecDelay(0.1)
    Call PowerSet2(1, "0.0", CurDevicePar.SetStdI1, 1, CurDevicePar.SetStdV2, CurDevicePar.SetStdI2, 1)
    Call MsecDelay(0.1)
    Call PowerSet2(0, CurDevicePar.SetStdV1, CurDevicePar.SetStdI1, 1, CurDevicePar.SetStdV2, CurDevicePar.SetStdI2, 1)
    If Dual_Flag Then
        cardresult = DO_WritePort(card, Channel_P1A, &HFE)
        Call MsecDelay(0.04)
        cardresult = DO_WritePort(card, Channel_P1A, &HFA)
    Else
        cardresult = DO_WritePort(card, Channel_P1A, &HFB)
    End If
    
    WaitDevOn ("058f")
    Call MsecDelay(0.2)
    MPFlag = 0
Else
    Call PowerSet2(1, "0.0", CurDevicePar.SetStdI1, 1, CurDevicePar.SetStdV2, CurDevicePar.SetStdI2, 1)
    Call MsecDelay(0.1)
    Call PowerSet2(0, CurDevicePar.SetStdV1, CurDevicePar.SetStdI1, 1, CurDevicePar.SetStdV2, CurDevicePar.SetStdI2, 1)
    If Dual_Flag Then
        cardresult = DO_WritePort(card, Channel_P1A, &HFE)
        Call MsecDelay(0.04)
        cardresult = DO_WritePort(card, Channel_P1A, &HFA)
    Else
        cardresult = DO_WritePort(card, Channel_P1A, &HFB)
    End If
    
    WaitDevOn ("058f")
    Call MsecDelay(0.2)
    
End If
        
    
OldTimer = Timer
AlcorMPMessage = 0
SetSiteStatus (RunHV)
MPTester.Print "RW Tester begin test........"
Call StartFTtest_AU69XX
    
Do
    If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
        AlcorMPMessage = mMsg.message
        TranslateMessage mMsg
        DispatchMessage mMsg
    End If
    
    PassTime = Timer - OldTimer
   
Loop Until AlcorMPMessage = WM_FT_RW_SPEED_FAIL _
        Or AlcorMPMessage = WM_FT_RW_RW_FAIL _
        Or AlcorMPMessage = WM_FT_RW_ROM_FAIL _
        Or AlcorMPMessage = WM_FT_RW_RAM_FAIL _
        Or AlcorMPMessage = WM_FT_RW_RW_PASS _
        Or AlcorMPMessage = WM_FT_RW_UNKNOW_FAIL _
        Or AlcorMPMessage = WM_FT_FLASH_NUM_FAIL _
        Or AlcorMPMessage = WM_FT_CHECK_CERBGPO_FAIL _
        Or AlcorMPMessage = WM_FT_CHECK_HW_CODE_FAIL _
        Or AlcorMPMessage = WM_FT_PHYREAD_FAIL _
        Or AlcorMPMessage = WM_FT_ECC_FAIL _
        Or AlcorMPMessage = WM_FT_NOFREEBLOCK_FAIL _
        Or AlcorMPMessage = WM_FT_LODECODE_FAIL _
        Or AlcorMPMessage = WM_FT_CHECK_WRITE_PROTECT_FAIL _
        Or AlcorMPMessage = WM_FT_NO_CARD_FAIL _
        Or AlcorMPMessage = WM_FT_TEST_DQS_FAIL Or AlcorMPMessage = WM_LC_FAIL Or AlcorMPMessage = WM_RC_FAIL Or PassTime > 10 _
        Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                
MPTester.Print "RW work Time="; PassTime
MPTester.MPText.Text = Hex(AlcorMPMessage)
            
'===========================================================
'  RW Time Out Fail
'===========================================================

If (PassTime > 10) Or ((FailCloseAP) And (AlcorMPMessage <> WM_FT_RW_RW_PASS)) Then
    Close_FT_AP ("UFD Test")
    
    If (PassTime > 10) Then
        TestResult = "Bin3"
        MPTester.TestResultLab = "Bin3:RW Time Out Fail"
        cardresult = DO_WritePort(card, Channel_P1A, &HFF)  ' power off
        SetSiteStatus (SiteUnknow)
        WaitAnotherSiteDone (HVDone)
        GoTo TestEnd
    End If
End If
                    
Select Case AlcorMPMessage
    Case WM_FT_RW_UNKNOW_FAIL
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2:UnKnow Fail"
        ContFail = ContFail + 1
    
    Case WM_FT_FLASH_NUM_FAIL
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2:Flash Number Fail"
        ContFail = ContFail + 1
    
    Case WM_FT_CHECK_HW_CODE_FAIL
        TestResult = "Bin5"
        MPTester.TestResultLab = "Bin5:HW-ID Fail"
        ContFail = ContFail + 1
    
'    Case WM_FT_TESTUNITREADY_FAIL
'        TestResult = "Bin2"
'        MPTester.TestResultLab = "Bin2:TestUnitReady Fail"
'        ContFail = ContFail + 1
    
    Case WM_FT_RW_SPEED_FAIL
        TestResult = "Bin3"
        MPTester.TestResultLab = "Bin3:SPEED Error "
        ContFail = ContFail + 1
         
    Case WM_FT_RW_RW_FAIL
        TestResult = "Bin3"
        MPTester.TestResultLab = "Bin3:RW FAIL "
        ContFail = ContFail + 1
    
    Case WM_FT_CHECK_CERBGPO_FAIL
        TestResult = "Bin3"
        MPTester.TestResultLab = "Bin3:GPO/RB FAIL "
        ContFail = ContFail + 1
    
    Case WM_FT_CHECK_WRITE_PROTECT_FAIL
        TestResult = "Bin3"
        MPTester.TestResultLab = "Bin3:W/P FAIL "
        ContFail = ContFail + 1
         
    Case WM_FT_NO_CARD_FAIL
        TestResult = "Bin4"
        MPTester.TestResultLab = "Bin4:NoCard FAIL "
        ContFail = ContFail + 1
    
    Case WM_FT_RW_ROM_FAIL
        TestResult = "Bin4"
        MPTester.TestResultLab = "Bin4:ROM FAIL "
        ContFail = ContFail + 1
          
    Case WM_FT_PHYREAD_FAIL
        TestResult = "Bin4"
        MPTester.TestResultLab = "Bin4:PHY Read FAIL "
        ContFail = ContFail + 1
          
    Case WM_FT_RW_RAM_FAIL
        TestResult = "Bin4"
        MPTester.TestResultLab = "Bin4:RAM FAIL "
        ContFail = ContFail + 1
           
    Case WM_FT_NOFREEBLOCK_FAIL
        TestResult = "Bin4"
        MPTester.TestResultLab = "Bin4:FreeBlock FAIL "
        ContFail = ContFail + 1
    
    Case WM_FT_LODECODE_FAIL
        TestResult = "Bin4"
        MPTester.TestResultLab = "Bin4:LoadCode FAIL "
        ContFail = ContFail + 1
    
'    Case WM_FT_RELOADCODE_FAIL
'        TestResult = "Bin4"
'        MPTester.TestResultLab = "Bin4:ReLoadCode FAIL "
'        ContFail = ContFail + 1
    
    Case WM_FT_ECC_FAIL
        TestResult = "Bin5"
        MPTester.TestResultLab = "Bin5:ECC FAIL "
        ContFail = ContFail + 1
        
    Case WM_FT_MOVE_DATA_FAIL
        TestResult = "Bin5"
        MPTester.TestResultLab = "Bin5:MOVE DATA FAIL "
        ContFail = ContFail + 1
                
    Case WM_FT_RW_RW_PASS
        If CheckLED_Flag Then
            For LedCount = 1 To 20
                Call MsecDelay(0.1)
                cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
                If LightOn = &HEF Or LightOn = &HCF Or LightOn = 223 Then
                    Exit For
                End If
            Next LedCount
            
            MPTester.Print "light="; LightOn
            
            If LightOn = &HEF Or LightOn = &HCF Or LightOn = 223 Then
                MPTester.TestResultLab = "PASS "
                TestResult = "PASS"
                ContFail = 0
            Else
                TestResult = "Bin3"
                MPTester.TestResultLab = "Bin3:LED FAIL "
                ContFail = ContFail + 1
            End If
        Else
            MPTester.TestResultLab = "PASS "
            TestResult = "PASS"
            ContFail = 0
        End If
        
    Case Else
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2:Undefine Fail"
        ContFail = ContFail + 1
           
End Select

SetSiteStatus (HVDone)
WaitAnotherSiteDone (HVDone)
    
TestEnd:
    
cardresult = DO_WritePort(card, Channel_P1A, &HFF)
Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
SetSiteStatus (SiteUnknow)
WaitDevOFF ("058f")
    
End Sub


