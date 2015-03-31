Attribute VB_Name = "AU7670"
Option Explicit

Public Const WM_USER = &H400
Public Const WM_TRIGGER_MP = WM_USER + &H1E
Public Const WM_THREAD_FTCHECK = WM_USER + &H1C
Public Const WM_THREAD_OK = WM_USER + &H1D
Public Const WM_THREAD_TERMINATE = WM_USER + &H1B

Public Const CARD_SHORT_ERROR = WM_USER + &H100
Public Const CHIP_VERSION_ERROR = WM_USER + &H201
Public Const GET_FLASH_INFO_ERROR_TP = WM_USER + &H400
Public Const UNKNOWN_FLASH_ERROR = WM_USER + &H403
Public Const LOAD_FIRMWARE_ERROR = WM_USER + &H800
Public Const HOST_OPEN_FIRMWARE_ERROR = WM_USER + &H900
Public Const NO_CARD_ERROR_TP = WM_USER + &HFFFF
Public Const SD_BUS_ERROR = WM_USER + &H150
Public Const SD_BUS_WRITE_ERROR_TP = WM_USER + &H151
Public Const SD_BUS_READ_ERROR_TP = WM_USER + &H152
Public Const FT_DATA_WRITE_ERROR = WM_USER + &H1654
Public Const FT_DATA_READ_ERROR = WM_USER + &H1655
Public Const FT_DATA_COMPARE_ERROR = WM_USER + &H1656
Public Const WRITE_FAT_ERROR_TP = WM_USER + &H1700
Public Const WRITE_FAT_COMPARE_DATA_ERROR = WM_USER + &H1701
Public Const CHECK_FLASH_ECC_ERROR_FT = WM_USER + &HA06

Public Const WM_PINCHECK_START = WM_USER + &H20
Public Const WM_PINCHECK_DONE = WM_USER + &H21
Public Const FT_SP_RAM_COMPARE_ERROR = WM_USER + &H1658
Public Const FT_DP_RAM_COMPARE_ERROR = WM_USER + &H1659
Public Const FT_RB_CHECK_ERROR = WM_USER + &H1657

Public Const GET_PIN_CHECK = WM_USER + &H22
Public TargetHwnd As Long

Public Sub AU7670XXXELF20TestSub()
Dim OldTimer, PassTime
Dim mMsg As MSG
Dim rt2

Dim WP_CE_CHECK(5) As Boolean
Dim WP_CE As Long

Dim WP_on As Boolean
Dim WP_off As Boolean
Dim CE0_on As Boolean
Dim CE0_off As Boolean
Dim CE1_on As Boolean
Dim CE1_off As Boolean
Dim CE2_on As Boolean
Dim CE2_off As Boolean
Dim CE3_on As Boolean
Dim CE3_off As Boolean
Dim Read_Port_Check As Boolean
Dim i As Integer


    MPTester.AU7670 = True     ' for 1ms delay mark if condition

    If MPTester.Check1.Value = 1 Then
        MPFlag = 0
    End If
    
    If PCI7248InitFinish = 0 Then
        PCI7248Exist
        SetTimer_1ms
    End If
    
    For i = 0 To 4
        WP_CE_CHECK(i) = False
    Next i
    
    cardresult = DO_WritePort(card, Channel_P1A, &H0)           ' for initail, bit 2 set to 0 for ENA on
    Call MsecDelay(0.1)
    winHwnd = FindWindow(vbNullString, "Alcor Micro SD MP")

    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\AU7670\SDMP.exe", "", "", SW_SHOW)
        MsecDelay (2)
        winHwnd = FindWindow(vbNullString, "Alcor Micro SD MP")
    End If
    
    SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags
    
    If MPFlag = 0 Or MPFlag = 5 Then
    
        If Not WaitDevOn("vid_058f") Then
            TestResult = "Bin2"
            MPTester.TestResultLab = "UnKnown Device"
            MPFlag = 1
            Exit Sub
        End If
                
        rt2 = PostMessage(winHwnd, WM_TRIGGER_MP, 0&, 0&)
        MsecDelay (0.5)
        
        OldTimer = Timer
            
        Do
            If PeekMessage(mMsg, 0, WM_USER, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
            
            PassTime = Timer - OldTimer

        Loop Until PassTime > 45 Or PassTime < 0 _
            Or AlcorMPMessage = CARD_SHORT_ERROR _
            Or AlcorMPMessage = CHIP_VERSION_ERROR _
            Or AlcorMPMessage = GET_FLASH_INFO_ERROR_TP _
            Or AlcorMPMessage = UNKNOWN_FLASH_ERROR _
            Or AlcorMPMessage = LOAD_FIRMWARE_ERROR _
            Or AlcorMPMessage = HOST_OPEN_FIRMWARE_ERROR _
            Or AlcorMPMessage = NO_CARD_ERROR_TP _
            Or AlcorMPMessage = SD_BUS_ERROR _
            Or AlcorMPMessage = SD_BUS_WRITE_ERROR_TP _
            Or AlcorMPMessage = SD_BUS_READ_ERROR_TP _
            Or AlcorMPMessage = FT_DATA_WRITE_ERROR _
            Or AlcorMPMessage = FT_DATA_READ_ERROR _
            Or AlcorMPMessage = FT_DATA_COMPARE_ERROR _
            Or AlcorMPMessage = WRITE_FAT_ERROR_TP _
            Or AlcorMPMessage = WRITE_FAT_COMPARE_DATA_ERROR _
            Or AlcorMPMessage = CHECK_FLASH_ECC_ERROR_FT _
            Or AlcorMPMessage = WM_THREAD_OK _
            Or AlcorMPMessage = FT_SP_RAM_COMPARE_ERROR _
            Or AlcorMPMessage = FT_DP_RAM_COMPARE_ERROR _
            Or AlcorMPMessage = FT_RB_CHECK_ERROR

        If PassTime > 45 Or PassTime < 0 Then
            MPTester.TestResultLab = "MP Time out Fail"
            TestResult = "Bin2"
            MPTester.Print "MP Time out Fail"
            MPFlag = 0
            cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        Else
            Call Binning("MP", AlcorMPMessage)
            AlcorMPMessage = 0
            TargetHwnd = mMsg.hwnd
        End If
    End If
    
    If MPFlag >= 1 Or MPFlag < 5 Then
    
        winHwnd = FindWindow(vbNullString, "Alcor Micro SD MP")
        
        If winHwnd = 0 Then
            Call ShellExecute(MPTester.hwnd, "open", App.Path & "\AU7670\SDMP.exe", "", "", SW_SHOW)
        End If
        
        SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags
        
        rt2 = PostMessage(winHwnd, WM_THREAD_FTCHECK, 0&, 0&)
        
        AlcorMPMessage = 0
        OldTimer = Timer
            
        Do
            If PeekMessage(mMsg, TargetHwnd, WM_USER, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
                DispatchMessage mMsg
            End If
            
            PassTime = Timer - OldTimer
            
        Loop Until PassTime > 20 Or PassTime < 0 _
                Or AlcorMPMessage = WM_PINCHECK_START _
                Or AlcorMPMessage = CARD_SHORT_ERROR _
                Or AlcorMPMessage = CHIP_VERSION_ERROR _
                Or AlcorMPMessage = GET_FLASH_INFO_ERROR_TP _
                Or AlcorMPMessage = UNKNOWN_FLASH_ERROR _
                Or AlcorMPMessage = LOAD_FIRMWARE_ERROR _
                Or AlcorMPMessage = HOST_OPEN_FIRMWARE_ERROR _
                Or AlcorMPMessage = NO_CARD_ERROR_TP _
                Or AlcorMPMessage = SD_BUS_ERROR _
                Or AlcorMPMessage = SD_BUS_WRITE_ERROR_TP _
                Or AlcorMPMessage = SD_BUS_READ_ERROR_TP _
                Or AlcorMPMessage = FT_DATA_WRITE_ERROR _
                Or AlcorMPMessage = FT_DATA_READ_ERROR _
                Or AlcorMPMessage = FT_DATA_COMPARE_ERROR _
                Or AlcorMPMessage = WRITE_FAT_ERROR_TP _
                Or AlcorMPMessage = WRITE_FAT_COMPARE_DATA_ERROR _
                Or AlcorMPMessage = CHECK_FLASH_ECC_ERROR_FT _
                Or AlcorMPMessage = FT_SP_RAM_COMPARE_ERROR _
                Or AlcorMPMessage = FT_DP_RAM_COMPARE_ERROR _
                Or AlcorMPMessage = FT_RB_CHECK_ERROR
                'Or AlcorMPMessage = WM_THREAD_OK
            
        If AlcorMPMessage = WM_PINCHECK_START And (PassTime < 20 Or PassTime < 0) Then
            ' for sending "RECEIVED PIN CHECK" MSG here
            MsecDelay (0.02)
            
            rt2 = PostMessage(winHwnd, GET_PIN_CHECK, 0&, 0&)
            MsecDelay (0.1)
        Else
            Call Binning("MP", AlcorMPMessage)
            MPFlag = MPFlag + 1
            rt2 = PostMessage(winHwnd, GET_PIN_CHECK, 0&, 0&)
            MsecDelay (0.1)
            GoTo ENDTEST
        End If
    
        OldTimer = Timer
    
        Do
            cardresult = DO_ReadPort(card, Channel_P1B, WP_CE)      ' read port2 for WP and CE pin

            If WP_CE_CHECK(0) = False Then
            
                If (WP_CE And &H1) = &H1 Then
                    WP_on = True
                Else
                    WP_off = True
                End If

                If WP_on And WP_off Then WP_CE_CHECK(0) = True
                
            End If

            If WP_CE_CHECK(1) = False Then
            
                If (WP_CE And &H2) = &H2 Then
                    CE0_on = True
                Else
                    CE0_off = True
                End If

                If CE0_on And CE0_off Then WP_CE_CHECK(1) = True

            End If

            If WP_CE_CHECK(2) = False Then
            
                If (WP_CE And &H4) = &H4 Then
                    CE1_on = True
                Else
                    CE1_off = True
                End If

                If CE1_on And CE1_off Then WP_CE_CHECK(2) = True

            End If

            If WP_CE_CHECK(3) = False Then
            
                If (WP_CE And &H8) = &H8 Then
                    CE2_on = True
                Else
                    CE2_off = True
                End If

                If CE2_on And CE2_off Then WP_CE_CHECK(3) = True

            End If

            If WP_CE_CHECK(4) = False Then
            
                If (WP_CE And &H10) = &H10 Then
                    CE3_on = True
                Else
                    CE3_off = True
                End If

                If CE3_on And CE3_off Then WP_CE_CHECK(4) = True
                
            End If

            If (WP_CE_CHECK(0) = True And WP_CE_CHECK(1) = True And WP_CE_CHECK(2) = True And WP_CE_CHECK(3) = True And WP_CE_CHECK(4) = True) Then
                Read_Port_Check = True
            Else
                Read_Port_Check = False
            End If

            PassTime = Timer - OldTimer

        Loop Until Read_Port_Check Or PassTime > 10 Or PassTime < 0
        
        If (Read_Port_Check = True) And PassTime < 10 Then
        
            OldTimer = Timer
            
            Do
                If Read_Port_Check = False Then
                    Exit Do
                Else
                    cardresult = DO_WritePort(card, Channel_P1A, &H8)      ' send hi lo signal
                    'Timer_1ms (1)
                    cardresult = DO_WritePort(card, Channel_P1A, &H0)
                    'Timer_1ms (1)
                    cardresult = DO_WritePort(card, Channel_P1A, &H10)       ' send hi lo signal
                    'Timer_1ms (1)
                    cardresult = DO_WritePort(card, Channel_P1A, &H0)
                    'Timer_1ms (1)
                End If
                
                If PeekMessage(mMsg, TargetHwnd, WM_USER, 0, PM_REMOVE) Then
                    AlcorMPMessage = mMsg.message
                    TranslateMessage mMsg
                    DispatchMessage mMsg
                End If
                
                PassTime = Timer - OldTimer
                
            ' when get stop msg or timeout then stop sending hi lo signal
            Loop Until PassTime > 15 Or PassTime < 0 _
                    Or AlcorMPMessage = WM_PINCHECK_DONE _
                    Or AlcorMPMessage = CARD_SHORT_ERROR _
                    Or AlcorMPMessage = SD_BUS_ERROR _
                    Or AlcorMPMessage = SD_BUS_WRITE_ERROR_TP _
                    Or AlcorMPMessage = SD_BUS_READ_ERROR_TP _
                    Or AlcorMPMessage = CHIP_VERSION_ERROR _
                    Or AlcorMPMessage = GET_FLASH_INFO_ERROR_TP _
                    Or AlcorMPMessage = UNKNOWN_FLASH_ERROR _
                    Or AlcorMPMessage = NO_CARD_ERROR_TP _
                    Or AlcorMPMessage = LOAD_FIRMWARE_ERROR _
                    Or AlcorMPMessage = HOST_OPEN_FIRMWARE_ERROR _
                    Or AlcorMPMessage = FT_DATA_WRITE_ERROR _
                    Or AlcorMPMessage = FT_DATA_READ_ERROR _
                    Or AlcorMPMessage = FT_DATA_COMPARE_ERROR _
                    Or AlcorMPMessage = WRITE_FAT_ERROR_TP _
                    Or AlcorMPMessage = WRITE_FAT_COMPARE_DATA_ERROR _
                    Or AlcorMPMessage = CHECK_FLASH_ECC_ERROR_FT _
                    Or AlcorMPMessage = FT_SP_RAM_COMPARE_ERROR _
                    Or AlcorMPMessage = FT_DP_RAM_COMPARE_ERROR _
                    Or AlcorMPMessage = FT_RB_CHECK_ERROR _
                    Or AlcorMPMessage = WM_THREAD_OK _
                    Or AlcorMPMessage = WM_THREAD_TERMINATE
        Else
            Call Binning("MP", AlcorMPMessage)
            MPFlag = MPFlag + 1
            rt2 = PostMessage(winHwnd, GET_PIN_CHECK, 0&, 0&)
            MsecDelay (0.1)
            GoTo ENDTEST
        End If
        
        If AlcorMPMessage = WM_PINCHECK_DONE And PassTime < 15 Then
            rt2 = PostMessage(winHwnd, GET_PIN_CHECK, 0&, 0&)
        Else
            Call Binning("MP", AlcorMPMessage)
            MPFlag = MPFlag + 1
            rt2 = PostMessage(winHwnd, GET_PIN_CHECK, 0&, 0&)
            MsecDelay (0.1)
            GoTo ENDTEST
        End If
        
        Do
            'DoEvents
            If PeekMessage(mMsg, TargetHwnd, WM_USER, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If

            PassTime = Timer - OldTimer
            
            Select Case AlcorMPMessage

                Case CARD_SHORT_ERROR, SD_BUS_ERROR, SD_BUS_WRITE_ERROR_TP, SD_BUS_READ_ERROR_TP
                    Exit Do
                Case CHIP_VERSION_ERROR, GET_FLASH_INFO_ERROR_TP, UNKNOWN_FLASH_ERROR, NO_CARD_ERROR_TP, LOAD_FIRMWARE_ERROR
                    Exit Do
                Case HOST_OPEN_FIRMWARE_ERROR, FT_DATA_WRITE_ERROR, FT_DATA_READ_ERROR, FT_DATA_COMPARE_ERROR, WRITE_FAT_ERROR_TP
                    Exit Do
                Case WRITE_FAT_COMPARE_DATA_ERROR, CHECK_FLASH_ECC_ERROR_FT, FT_SP_RAM_COMPARE_ERROR, FT_DP_RAM_COMPARE_ERROR, FT_RB_CHECK_ERROR
                    Exit Do
                Case WM_THREAD_OK, WM_THREAD_TERMINATE
                    Exit Do
            End Select
            
        Loop Until PassTime > 30 Or PassTime < 0

        If PassTime > 30 Or PassTime < 0 Then
            MPTester.TestResultLab = "FT Fail"
            TestResult = "Bin2"
            MPFlag = MPFlag + 1
            MPTester.Print "FT Fail"
        Else
            Call Binning("FT", AlcorMPMessage)
        End If
    End If

ENDTEST:

    If TestResult <> "PASS" Then
        If MPFlag > 5 Then
            Close_AU7670_AP
            MPFlag = 0
        End If
    End If
    
End Sub

Private Sub Binning(MP_FT As String, MSG As Long)

Dim rt2

    If MP_FT = "MP" Then
        If MSG = CARD_SHORT_ERROR Then
            MPTester.TestResultLab = "Card Short Error"
            TestResult = "Bin2"
            MPTester.Print "Card Short Error"
            MPFlag = 0
            GoTo Dev_off
        ElseIf MSG = CHIP_VERSION_ERROR Then
            MPTester.TestResultLab = "Chip Version Error"
            TestResult = "Bin3"
            MPTester.Print "Chip Version Error"
            MPFlag = 0
            GoTo Dev_off
        ElseIf MSG = GET_FLASH_INFO_ERROR_TP Then
            MPTester.TestResultLab = "Get Flash Info Error"
            TestResult = "Bin3"
            MPTester.Print "Get Flash Info Error"
            MPFlag = 0
            GoTo Dev_off
        ElseIf MSG = UNKNOWN_FLASH_ERROR Then
            MPTester.TestResultLab = "Unknown Flash Error"
            TestResult = "Bin3"
            MPTester.Print "Unknown Flash Error"
            MPFlag = 0
            GoTo Dev_off
        ElseIf MSG = LOAD_FIRMWARE_ERROR Then
            MPTester.TestResultLab = "Load Firmware Error"
            TestResult = "Bin4"
            MPTester.Print "Load Firmware Error"
            MPFlag = 0
            GoTo Dev_off
        ElseIf MSG = HOST_OPEN_FIRMWARE_ERROR Then
            MPTester.TestResultLab = "Host Open Firmware Error"
            TestResult = "Bin2"
            MPTester.Print "Host Open Firmware Error"
            MPFlag = 0
            GoTo Dev_off
        ElseIf MSG = NO_CARD_ERROR_TP Then
            MPTester.TestResultLab = "No Card Error"
            TestResult = "Bin3"
            MPTester.Print "No Card Error"
            MPFlag = 0
            GoTo Dev_off
        ElseIf MSG = SD_BUS_ERROR Then
            MPTester.TestResultLab = "SD Bus Error"
            TestResult = "Bin2"
            MPTester.Print "SD Bus Error"
            MPFlag = 0
            GoTo Dev_off
        ElseIf MSG = SD_BUS_WRITE_ERROR_TP Then
            MPTester.TestResultLab = "SD Bus Write Error"
            TestResult = "Bin2"
            MPTester.Print "SD Bus Write Error"
            MPFlag = 0
            GoTo Dev_off
        ElseIf MSG = SD_BUS_READ_ERROR_TP Then
            MPTester.TestResultLab = "SD Bus Read Error"
            TestResult = "Bin2"
            MPTester.Print "SD Bus Read Error"
            MPFlag = 0
            GoTo Dev_off
        ElseIf MSG = FT_DATA_WRITE_ERROR Then
            MPTester.TestResultLab = "FT Data Write Error"
            TestResult = "Bin5"
            MPTester.Print "FT Data Write Error"
            MPFlag = 0
            GoTo Dev_off
        ElseIf MSG = FT_DATA_READ_ERROR Then
            MPTester.TestResultLab = "FT Data Read Error"
            TestResult = "Bin5"
            MPTester.Print "FT Data Read Error"
            MPFlag = 0
            GoTo Dev_off
        ElseIf MSG = FT_DATA_COMPARE_ERROR Then
            MPTester.TestResultLab = "FT Data Compare Error"
            TestResult = "Bin5"
            MPTester.Print "FT Data Compare Error"
            MPFlag = 0
            GoTo Dev_off
        ElseIf MSG = WRITE_FAT_ERROR_TP Then
            MPTester.TestResultLab = "Write FAT Error"
            TestResult = "Bin5"
            MPTester.Print "Write FAT Error"
            MPFlag = 0
            GoTo Dev_off
        ElseIf MSG = WRITE_FAT_COMPARE_DATA_ERROR Then
            MPTester.TestResultLab = "Write FAT Compare Data Error"
            TestResult = "Bin5"
            MPTester.Print "Write FAT Compare Data Error"
            MPFlag = 0
            GoTo Dev_off
        ElseIf MSG = CHECK_FLASH_ECC_ERROR_FT Then
            MPTester.TestResultLab = "Check Flash ECC Error"
            TestResult = "Bin5"
            MPTester.Print "Check Flash ECC Error"
            MPFlag = 0
            GoTo Dev_off
        ElseIf MSG = FT_SP_RAM_COMPARE_ERROR Then
            MPTester.TestResultLab = "FT SP Ram Compare Error"
            TestResult = "Bin4"
            MPTester.Print "FT SP Ram Compare Error"
            MPFlag = 0
            GoTo Dev_off
        ElseIf MSG = FT_DP_RAM_COMPARE_ERROR Then
            MPTester.TestResultLab = "FT DP Ram Compare Error"
            TestResult = "Bin4"
            MPTester.Print "FT DP Ram Compare Error"
            MPFlag = 0
            GoTo Dev_off
        ElseIf MSG = FT_RB_CHECK_ERROR Then
            MPTester.TestResultLab = "FT RB Check Error"
            TestResult = "Bin5"
            MPTester.Print "FT RB Check Error"
            MPFlag = 0
            GoTo Dev_off
        ElseIf MSG = WM_THREAD_OK Then
        
            rt2 = PostMessage(winHwnd, GET_PIN_CHECK, 0&, 0&)
        
            MPTester.TestResultLab = "MP PASS"
            MPFlag = 1
            Exit Sub
'        ElseIf MSG = WM_THREAD_TERMINATE Then
'            MPTester.TestResultLab = "MP Ready Fail"
'            TestResult = "Bin2"
'            MPTester.Print "MP Ready Fail"
'            MPFlag = 0
'            GoTo Dev_off
        End If
    Else
'        If AlcorMPMessage = WM_THREAD_TERMINATE Then
'            MPTester.TestResultLab = "FT Fail"
'            TestResult = "Bin2"
'            MPFlag = MPFlag + 1
'            MPTester.Print "FT Fail"
'        Else
        If AlcorMPMessage = FT_SP_RAM_COMPARE_ERROR Then
            MPTester.TestResultLab = "FT SP Ram Compare Error"
            TestResult = "Bin4"
            MPTester.Print "FT SP Ram Compare Error"
            MPFlag = MPFlag + 1
        ElseIf AlcorMPMessage = FT_DP_RAM_COMPARE_ERROR Then
            MPTester.TestResultLab = "FT DP Ram Compare Error"
            TestResult = "Bin4"
            MPTester.Print "FT DP Ram Compare Error"
            MPFlag = MPFlag + 1
        ElseIf AlcorMPMessage = CARD_SHORT_ERROR Then
            MPTester.TestResultLab = "Card Short Error"
            TestResult = "Bin2"
            MPFlag = MPFlag + 1
            MPTester.Print "Card Short Error"
        ElseIf AlcorMPMessage = CHIP_VERSION_ERROR Then
            MPTester.TestResultLab = "Chip Version Error"
            TestResult = "Bin3"
            MPFlag = MPFlag + 1
            MPTester.Print "Chip Version Error"
        ElseIf AlcorMPMessage = GET_FLASH_INFO_ERROR_TP Then
            MPTester.TestResultLab = "Get Flash Info Error"
            TestResult = "Bin3"
            MPFlag = MPFlag + 1
            MPTester.Print "Get Flash Info Error"
        ElseIf AlcorMPMessage = UNKNOWN_FLASH_ERROR Then
            MPTester.TestResultLab = "Unknown Flash Error"
            TestResult = "Bin3"
            MPFlag = MPFlag + 1
            MPTester.Print "Unknown Flash Error"
        ElseIf AlcorMPMessage = LOAD_FIRMWARE_ERROR Then
            MPTester.TestResultLab = "Load Firmware Error"
            TestResult = "Bin4"
            MPFlag = MPFlag + 1
            MPTester.Print "Load Firmware Error"
        ElseIf AlcorMPMessage = HOST_OPEN_FIRMWARE_ERROR Then
            MPTester.TestResultLab = "Host Open Firmware Error"
            TestResult = "Bin2"
            MPFlag = MPFlag + 1
            MPTester.Print "Host Open Firmware Error"
        ElseIf AlcorMPMessage = NO_CARD_ERROR_TP Then
            MPTester.TestResultLab = "No Card Error"
            TestResult = "Bin3"
            MPFlag = MPFlag + 1
            MPTester.Print "No Card Error"
        ElseIf AlcorMPMessage = SD_BUS_ERROR Then
            MPTester.TestResultLab = "SD Bus Error"
            TestResult = "Bin2"
            MPFlag = MPFlag + 1
            MPTester.Print "SD Bus Error"
        ElseIf AlcorMPMessage = SD_BUS_WRITE_ERROR_TP Then
            MPTester.TestResultLab = "SD Bus Write Error"
            TestResult = "Bin2"
            MPFlag = MPFlag + 1
            MPTester.Print "SD Bus Write Error"
        ElseIf AlcorMPMessage = SD_BUS_READ_ERROR_TP Then
            MPTester.TestResultLab = "SD Bus Read Error"
            TestResult = "Bin2"
            MPFlag = MPFlag + 1
            MPTester.Print "SD Bus Read Error"
        ElseIf AlcorMPMessage = FT_DATA_WRITE_ERROR Then
            MPTester.TestResultLab = "FT Data Write Error"
            TestResult = "Bin5"
            MPFlag = MPFlag + 1
            MPTester.Print "FT Data Write Error"
        ElseIf AlcorMPMessage = FT_DATA_READ_ERROR Then
            MPTester.TestResultLab = "FT Data Read Error"
            TestResult = "Bin5"
            MPFlag = MPFlag + 1
            MPTester.Print "FT Data Read Error"
        ElseIf AlcorMPMessage = FT_DATA_COMPARE_ERROR Then
            MPTester.TestResultLab = "FT Data Compare Error"
            TestResult = "Bin5"
            MPFlag = MPFlag + 1
            MPTester.Print "FT Data Compare Error"
        ElseIf AlcorMPMessage = WRITE_FAT_ERROR_TP Then
            MPTester.TestResultLab = "Write FAT Error"
            TestResult = "Bin5"
            MPFlag = MPFlag + 1
            MPTester.Print "Write FAT Error"
        ElseIf AlcorMPMessage = WRITE_FAT_COMPARE_DATA_ERROR Then
            MPTester.TestResultLab = "Write FAT Compare Data Error"
            TestResult = "Bin5"
            MPFlag = MPFlag + 1
            MPTester.Print "Write FAT Compare Data Error"
        ElseIf AlcorMPMessage = CHECK_FLASH_ECC_ERROR_FT Then
            MPTester.TestResultLab = "Check Flash ECC Error"
            TestResult = "Bin5"
            MPFlag = MPFlag + 1
            MPTester.Print "Check Flash ECC Error"
        ElseIf AlcorMPMessage = WM_THREAD_OK Then
        
            rt2 = PostMessage(winHwnd, GET_PIN_CHECK, 0&, 0&)
        
            MPTester.TestResultLab = "PASS"
            TestResult = "PASS"
            MPTester.Print "PASS"
        Else
            MPTester.TestResultLab = "Unknown Fail"
            TestResult = "Bin2"
            MPFlag = MPFlag + 1
            MPTester.Print "Unknown Fail"
        End If
    End If
    
Dev_off:
    cardresult = DO_WritePort(card, Channel_P1A, &HFF)

End Sub

Public Sub Close_AU7670_AP()
Dim rt2 As Long
Dim EntryTime As Long
Dim PassingTime As Long
Dim mMsg As MSG
Dim winChildHwnd As Long


    winHwnd = FindWindow(vbNullString, "Alcor Micro SD MP")
    
    EntryTime = Timer
    
    If winHwnd <> 0 Then
        Do
            rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
            Call MsecDelay(0.3)
            
            Do
                winChildHwnd = FindWindow(vbNullString, "SDMP")
            Loop Until (winChildHwnd <> 0) Or (Timer - EntryTime > 3)
            
            If (winChildHwnd <> 0) Then
                rt2 = PostMessage(winChildHwnd, WM_QUIT, 0&, 0&)
            End If
        
            winHwnd = FindWindow(vbNullString, "Alcor Micro SD MP")
            Call MsecDelay(0.5)
            
            PassingTime = Timer - EntryTime
        Loop Until (winHwnd <> 0) Or (PassingTime >= 6)
    End If
    
End Sub
