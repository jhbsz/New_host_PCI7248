Attribute VB_Name = "AU3821CAM"
Option Explicit

Public Type SYSTEM_INFO
    dwOemID As Long
    wProcessorArchitecture As Long
    wReserved As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Integer
    wProcessLevel As Integer
    wProcessorRevision As Integer
End Type

Public Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO) 'Get CPU Series Number

Public CPUInfo As SYSTEM_INFO

Public Const WM_USER = &H400
Public Const WM_CAM_MP_START = WM_USER + &H100
Public Const WM_CAM_MP_PASS = WM_USER + &H110
Public Const WM_CAM_MP_UNKNOW_FAIL = WM_USER + &H120
Public Const WM_CAM_MP_GPIO_FAIL = WM_USER + &H130

Public Const WM_CAM_SC_START = WM_USER + &H200
Public Const WM_CAM_SC_PASS = WM_USER + &H210
Public Const WM_CAM_SC_FAIL = WM_USER + &H220

Public Const WM_CAM_CSET_START = WM_USER + &H300
Public Const WM_CAM_CSET_PASS = WM_USER + &H310
Public Const WM_CAM_CSET_FAIL = WM_USER + &H320

' ========= 20130417 add for AU3826 ==========
Public Const WM_CAM_MP_MIPI_START = WM_USER + &H100            ' image pattern test(mipi)
Public Const WM_CAM_MP_PARA_START = WM_USER + &H101            ' image pattern test(parallel)
' ============================================


'########################################################
' 2012/5/31 for AU3825A61-DAF/DBF ST2 (Needs Win7 OS)
'########################################################
Public Const WM_CAM_PHY_START = WM_USER + &H600                '//PHY suspend/resume TEST
Public Const WM_CAM_PHY_PASS = WM_USER + &H610                 '//PHY suspend/resume TEST
Public Const WM_CAM_PHY_UNKNOW_FAIL = WM_USER + &H620          '//PHY suspend/resume TEST
Public Const WM_CAM_PHY_FAIL = WM_USER + &H630                 '//PHY suspend/resume TEST
'########################################################

Public Const WM_CAM_GPIO_SETTING = WM_USER + &H400
Public Const WM_CAM_GPIO_READ = WM_USER + &H410

Public Const WM_CAM_MP_READY = WM_USER + &H800

Public Const WM_CAM_FWCHECK_FAIL = WM_USER + &H720              'FW lost

'########################################################
' 2012/6/12 for AU3825A61-DAF/DBF ST4 (Phy-Board test)
'########################################################
'config P1C H-byte as input     (H-byte Bit1: Comp0(GPIO0), bit2: Comp1(GPIO1))
'config P1C L-byte as output    (L-byte Bit1: Reset(GPIO4))

Public Const PHY_TEST_PASS = &HC
Public Const PHY_TEST_Fail_1 = &HE
Public Const PHY_TEST_Fail_2 = &HD
Public Const PHY_TEST_UNKNOW = &HF

'########################################################



Public TempCounter As Integer
Public TempSRAMCounter As Integer
Public TempCSETCounter As Integer
Public TempGPIOCounter As Integer
Public TempConditionCounter As Integer
Public TempPASSCounter As Integer
Public TempImageCounter As Integer
Public TempLDOCounter As Integer
Public TempVDD18Counter As Integer
Public DriverDieCount As Integer
Public FW_Fail_Flag As Boolean

Public LDOVal() As Long
Public Sub SetTimer_1ms()
Dim ERR As Integer
ERR = CTR_Setup(card, 1, 2, 200, 0)
ERR = CTR_Setup(card, 2, 2, 10, 0)
End Sub
Public Sub SetTimer_500us()
Dim ERR As Integer
ERR = CTR_Setup(card, 1, 2, 200, 0)
ERR = CTR_Setup(card, 2, 2, 5, 0)
End Sub
Public Sub Timer_1ms(ms As Integer)
Dim result
Dim old_value1
Dim old_value2
Dim i As Integer
Dim ii As Integer
Dim T1 As Long
Dim T2 As Long

 
result = CTR_Read(0, 2, old_value1)

T1 = Timer
   For i = 1 To ms
   
   
            Do
            DoEvents
            result = CTR_Read(0, 2, old_value2)
            T2 = Timer
                If T2 - T1 > 1 Then
                    MsgBox ("PCI7248_Counter Error !!")
                    End
                End If
            Loop Until old_value1 <> old_value2
    
            Do
            DoEvents
            result = CTR_Read(0, 2, old_value2)
            T2 = Timer
                If T2 - T1 > 1 Then
                    MsgBox ("PCI7248_Counter Error !!")
                    End
                End If
            Loop Until old_value1 = old_value2
    
    
    
    Next
     
End Sub
Public Sub Timer_500us(us As Integer)
Dim result
Dim old_value1
Dim old_value2
Dim i As Integer
Dim ii As Integer
Dim T1 As Long
Dim T2 As Long

 
result = CTR_Read(0, 2, old_value1)

T1 = Timer
    For i = 1 To us
        Do
            DoEvents
            result = CTR_Read(0, 2, old_value2)
            T2 = Timer
                If T2 - T1 > 1 Then
                    MsgBox ("PCI7248_Counter Error !!")
                    End
                End If
            Loop Until old_value1 <> old_value2
            
            Do
            DoEvents
            result = CTR_Read(0, 2, old_value2)
            T2 = Timer
                If T2 - T1 > 1 Then
                    MsgBox ("PCI7248_Counter Error !!")
                    End
                End If
        Loop Until old_value1 = old_value2
    Next
     
End Sub

Public Sub LoadMP_Click_AU3821()

Dim TimePass
Dim rt2
    ' find window
    winHwnd = FindWindow(vbNullString, "VideoCap")
 
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\CamTest\" & ChipName & "\VideoCap.exe", "", "", SW_SHOW)
    End If

    SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags

End Sub

Public Function LoadVedioCap_AU3826A81FTTest_40QFN() As Boolean
Dim rt2
Dim TempPassingTime As Long
Dim TempOldTime As Long
Dim mMsg As MSG

     ' find window
    winHwnd = FindWindow(vbNullString, "VideoCap")
    LoadVedioCap_AU3826A81FTTest_40QFN = False
    
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\CamTest\AU3826A81FTTest_40QFN\VideoCap.exe", "", "", SW_SHOW)
    Else
        LoadVedioCap_AU3826A81FTTest_40QFN = True
        Exit Function
    End If
    
    SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags
    AlcorMPMessage = 0
    
    TempOldTime = Timer
    
    Do
        ' DoEvents
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
        
        TempPassingTime = Timer - TempOldTime
    
    Loop Until AlcorMPMessage = WM_CAM_MP_READY Or TempPassingTime > 10

    If AlcorMPMessage = WM_CAM_MP_READY Then
        LoadVedioCap_AU3826A81FTTest_40QFN = True
    End If

End Function

Public Function LoadVedioCap_AU3825_40QFN() As Boolean
Dim rt2
Dim TempPassingTime As Long
Dim TempOldTime As Long
Dim mMsg As MSG

     ' find window
    winHwnd = FindWindow(vbNullString, "VideoCap")
    LoadVedioCap_AU3825_40QFN = False
    
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\CamTest\AU3825A61FTTest_40QFN\VideoCap.exe", "", "", SW_SHOW)
    Else
        LoadVedioCap_AU3825_40QFN = True
        Exit Function
    End If
    
    SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags
    AlcorMPMessage = 0
    
    TempOldTime = Timer
    
    Do
        ' DoEvents
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
        
        TempPassingTime = Timer - TempOldTime
    
    Loop Until AlcorMPMessage = WM_CAM_MP_READY Or TempPassingTime > 10

    If AlcorMPMessage = WM_CAM_MP_READY Then
        LoadVedioCap_AU3825_40QFN = True
    End If

End Function

Public Function LoadVedioCap_AU3825_28QFN() As Boolean
Dim rt2
Dim TempPassingTime As Long
Dim TempOldTime As Long
Dim mMsg As MSG

     ' find window
    winHwnd = FindWindow(vbNullString, "VideoCap")
    LoadVedioCap_AU3825_28QFN = False
    
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\CamTest\AU3825A61FTTest_28QFN\VideoCap.exe", "", "", SW_SHOW)
    Else
        LoadVedioCap_AU3825_28QFN = True
        Exit Function
    End If
    
    SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags
    AlcorMPMessage = 0
    
    TempOldTime = Timer
    
    Do
        ' DoEvents
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
        
        TempPassingTime = Timer - TempOldTime
    
    Loop Until AlcorMPMessage = WM_CAM_MP_READY Or TempPassingTime > 10

    If AlcorMPMessage = WM_CAM_MP_READY Then
        LoadVedioCap_AU3825_28QFN = True
    End If

End Function

Public Function LoadVedioCap_AU3821() As Boolean
Dim rt2
Dim TempPassingTime As Long
Dim TempOldTime As Long
Dim mMsg As MSG

     ' find window
    winHwnd = FindWindow(vbNullString, "VideoCap")
    LoadVedioCap_AU3821 = False
    
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\CamTest\AU3821A66FNF21\VideoCap.exe", "", "", SW_SHOW)
    Else
        LoadVedioCap_AU3821 = True
        Exit Function
    End If
    
    SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags
    AlcorMPMessage = 0
    
    TempOldTime = Timer
    
    Do
        ' DoEvents
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
        
        TempPassingTime = Timer - TempOldTime
    
    Loop Until AlcorMPMessage = WM_CAM_MP_READY Or TempPassingTime > 10

    If AlcorMPMessage = WM_CAM_MP_READY Then
        LoadVedioCap_AU3821 = True
    End If

End Function
Public Sub Load_VerifyFW_Tool_AU3826A81FTTest_40QFN()

Dim TimePass
Dim rt2
Dim OldPath As String
Dim StatusStr As String
Dim OldTime As Long

    OldPath = CurDir
    ChDir (App.Path & "\CamTest\AU3826A81FTTest_40QFN\FW update Tool\")
    ' find window
    winHwnd = FindWindow(vbNullString, "AlcorMPTool v3.2.4")
    
    If Dir(App.Path & "\CamTest\AU3826A81FTTest_40QFN\FW update Tool\status.txt") = "status.txt" Then
        Kill (App.Path & "\CamTest\AU3826A81FTTest_40QFN\FW update Tool\status.txt")
    End If
    
    OldTime = Timer
    
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\CamTest\AU3826A81FTTest_40QFN\FW update Tool\VerifyFW_v3.2.4.exe", "", "", SW_SHOW)
    End If
    
    Call MsecDelay(1#)
    
    Do
        winHwnd = FindWindow(vbNullString, "AlcorMPTool v3.2.4")
        Call MsecDelay(0.5)
    Loop While (winHwnd <> 0) And (Timer - OldTime < 10)
    
    Call MsecDelay(0.4)
    
    If Dir(App.Path & "\CamTest\AU3826A81FTTest_40QFN\FW update Tool\status.txt") = "status.txt" Then
        Open App.Path & "\CamTest\AU3826A81FTTest_40QFN\FW update Tool\status.txt" For Input As #20
        Line Input #20, StatusStr
        
        If (InStr(StatusStr, "Fail") > 0) Or (InStr(StatusStr, "checking") > 0) Then
            FW_Fail_Flag = True
            
            ' Auto update FW
            cardresult = DO_WritePort(card, Channel_P1A, &HF6)  ' disable WP
            Call ShellExecute(MPTester.hwnd, "open", App.Path & "\CamTest\AU3826A81FTTest_40QFN\FW update Tool\MPTool_lite_v3.12.620.exe", "", "", SW_SHOW)
            Call MsecDelay(5#)
            
            Do
                winHwnd = FindWindow(vbNullString, "MPTool_lite_v3.12.620")
                Call MsecDelay(0.5)
            Loop While (winHwnd <> 0) And (Timer - OldTime < 10)
            
            Call MsecDelay(0.4)
            
            'End
        End If
    
        Close #20
    Else
        FW_Fail_Flag = False
    End If
    
    ChDir (OldPath)
End Sub

Public Sub Load_VerifyFW_Tool_AU3825_40QFN()

Dim TimePass
Dim rt2
Dim OldPath As String
Dim StatusStr As String
Dim OldTime As Long

    OldPath = CurDir
    ChDir (App.Path & "\CamTest\AU3825A61FTTest_40QFN\FW update Tool\")
    ' find window
    winHwnd = FindWindow(vbNullString, "AlcorMPTool v3.2.4")
    
    If Dir(App.Path & "\CamTest\AU3825A61FTTest_40QFN\FW update Tool\status.txt") = "status.txt" Then
        Kill (App.Path & "\CamTest\AU3825A61FTTest_40QFN\FW update Tool\status.txt")
    End If
    
    OldTime = Timer
    
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\CamTest\AU3825A61FTTest_40QFN\FW update Tool\VerifyFW_v3.2.4.exe", "", "", SW_SHOW)
    End If
    
    Call MsecDelay(1#)
    
    Do
        winHwnd = FindWindow(vbNullString, "AlcorMPTool v3.2.4")
        Call MsecDelay(0.5)
    Loop While (winHwnd <> 0) And (Timer - OldTime < 10)
    
    Call MsecDelay(0.4)
    
    If Dir(App.Path & "\CamTest\AU3825A61FTTest_40QFN\FW update Tool\status.txt") = "status.txt" Then
        Open App.Path & "\CamTest\AU3825A61FTTest_40QFN\FW update Tool\status.txt" For Input As #20
        Line Input #20, StatusStr
        
        If (InStr(StatusStr, "Fail") > 0) Or (InStr(StatusStr, "checking") > 0) Then
            FW_Fail_Flag = True
            MsgBox ("Please Manul Update FW Version")
            End
        End If
    
        Close #20
    Else
        FW_Fail_Flag = False
    End If
    
    ChDir (OldPath)
End Sub

Public Sub Load_VerifyFW_Tool_AU3825_28QFN()

Dim TimePass
Dim rt2
Dim OldPath As String
Dim StatusStr As String
Dim OldTime As Long

    OldPath = CurDir
    ChDir (App.Path & "\CamTest\AU3825A61FTTest_28QFN\FW update Tool\")
    ' find window
    winHwnd = FindWindow(vbNullString, "AlcorMPTool v3.2.4")
 
    If Dir(App.Path & "\CamTest\AU3825A61FTTest_28QFN\FW update Tool\status.txt") = "status.txt" Then
        Kill (App.Path & "\CamTest\AU3825A61FTTest_28QFN\FW update Tool\status.txt")
    End If
    
    OldTime = Timer
    
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\CamTest\AU3825A61FTTest_28QFN\FW update Tool\VerifyFW_v3.2.4.exe", "", "", SW_SHOW)
    End If
    
    Call MsecDelay(1#)
    
    Do
        winHwnd = FindWindow(vbNullString, "AlcorMPTool v3.2.4")
        Call MsecDelay(0.5)
    Loop While (winHwnd <> 0) And (Timer - OldTime < 10)
    
    Call MsecDelay(0.4)
    
    If Dir(App.Path & "\CamTest\AU3825A61FTTest_28QFN\FW update Tool\status.txt") = "status.txt" Then
        Open App.Path & "\CamTest\AU3825A61FTTest_28QFN\FW update Tool\status.txt" For Input As #20
        Line Input #20, StatusStr
        
        Call MsecDelay(0.2)
        If (InStr(StatusStr, "Fail") > 0) Or (InStr(StatusStr, "checking") > 0) Then
            FW_Fail_Flag = True
            MsgBox ("Please Manul Update FW Version")
            End
        End If
        
        Close #20
    Else
        FW_Fail_Flag = False
    End If
    
    ChDir (OldPath)
End Sub
Public Function Check_Parallel() As Byte
Dim rt2
Dim TempPassingTime As Long
Dim TempOldTime As Long
Dim mMsg As MSG

    Check_Parallel = 0
    
    winHwnd = FindWindow(vbNullString, "VideoCap")
    If winHwnd = 0 Then
        Check_Parallel = 0
        Exit Function
    End If
    
    TempOldTime = Timer
    rt2 = PostMessage(winHwnd, WM_CAM_MP_PARA_START, 0&, 0&)
 
    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
         
        TempPassingTime = Timer - TempOldTime
       
    Loop Until AlcorMPMessage = WM_CAM_MP_PASS _
          Or AlcorMPMessage = WM_CAM_MP_UNKNOW_FAIL _
          Or AlcorMPMessage = WM_CAM_MP_GPIO_FAIL _
          Or TempPassingTime > 10
          
    
    If AlcorMPMessage = WM_CAM_MP_PASS Then
        Check_Parallel = 1
    Else
        Check_Parallel = 0
    End If
    
    
End Function

Public Function Check_MIPI() As Byte
Dim rt2
Dim TempPassingTime As Long
Dim TempOldTime As Long
Dim mMsg As MSG

    Check_MIPI = 0
    
    winHwnd = FindWindow(vbNullString, "VideoCap")
    If winHwnd = 0 Then
        Check_MIPI = 0
        Exit Function
    End If
    
    TempOldTime = Timer
    rt2 = PostMessage(winHwnd, WM_CAM_MP_MIPI_START, 0&, 0&)
 
    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
         
        TempPassingTime = Timer - TempOldTime
       
    Loop Until AlcorMPMessage = WM_CAM_MP_PASS _
          Or AlcorMPMessage = WM_CAM_MP_UNKNOW_FAIL _
          Or AlcorMPMessage = WM_CAM_MP_GPIO_FAIL _
          Or TempPassingTime > 10
          'Or AlcorMPMessage = WM_CAM_CSET_PASS _
          'Or AlcorMPMessage = WM_CAM_FWCHECK_FAIL _

    
    If AlcorMPMessage = WM_CAM_MP_PASS Then
        Check_MIPI = 1
    Else
        Check_MIPI = 0
    End If
    
'    If AlcorMPMessage = WM_CAM_FWCHECK_FAIL Then
'        FW_Fail_Flag = True
'    End If
    
End Function

Public Sub StartRWTest_Click_AU3821()

Dim rt2
    winHwnd = FindWindow(vbNullString, "VideoCap")
    rt2 = PostMessage(winHwnd, WM_CAM_MP_START, 0&, 0&)

End Sub

Public Function CSET_Value_Test() As Byte
Dim rt2
Dim TempPassingTime As Long
Dim TempOldTime As Long
Dim mMsg As MSG

    CSET_Value_Test = 0
    
    winHwnd = FindWindow(vbNullString, "VideoCap")
    If winHwnd = 0 Then
        CSET_Value_Test = 0
        Exit Function
    End If
    
    TempOldTime = Timer
    rt2 = PostMessage(winHwnd, WM_CAM_CSET_START, 0&, 0&)
 
    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
         
        TempPassingTime = Timer - TempOldTime
       
    Loop Until AlcorMPMessage = WM_CAM_CSET_PASS _
          Or AlcorMPMessage = WM_CAM_CSET_FAIL _
          Or AlcorMPMessage = WM_CAM_FWCHECK_FAIL _
          Or TempPassingTime > 2
    

    If AlcorMPMessage = WM_CAM_CSET_PASS Then
        CSET_Value_Test = 1
    Else
        CSET_Value_Test = 0
    End If
    
    If AlcorMPMessage = WM_CAM_FWCHECK_FAIL Then
        FW_Fail_Flag = True
    End If
    
End Function

Public Function SRAM_Test() As Byte
Dim rt2
Dim TempPassingTime As Long
Dim TempOldTime As Long
Dim mMsg As MSG

    SRAM_Test = 0
    
    winHwnd = FindWindow(vbNullString, "VideoCap")
    If winHwnd = 0 Then
        SRAM_Test = 0
        Exit Function
    End If
    
    TempOldTime = Timer
    rt2 = PostMessage(winHwnd, WM_CAM_SC_START, 0&, 0&)
 
    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
         
        TempPassingTime = Timer - TempOldTime
       
    Loop Until AlcorMPMessage = WM_CAM_SC_PASS _
          Or AlcorMPMessage = WM_CAM_SC_FAIL _
          Or AlcorMPMessage = WM_CAM_FWCHECK_FAIL _
          Or TempPassingTime > 5

    If AlcorMPMessage = WM_CAM_SC_PASS Then
        SRAM_Test = 1
    Else
        SRAM_Test = 0
    End If
    
    If AlcorMPMessage = WM_CAM_FWCHECK_FAIL Then
        FW_Fail_Flag = True
    End If

End Function

Public Function Image_Test() As Byte
Dim rt2
Dim TempPassingTime As Long
Dim TempOldTime As Long
Dim mMsg As MSG

    Image_Test = 0
    TempPassingTime = 0
    
    winHwnd = FindWindow(vbNullString, "VideoCap")
    If winHwnd = 0 Then
        Image_Test = 0
        Exit Function
    End If
    
    TempOldTime = Timer
    rt2 = PostMessage(winHwnd, WM_CAM_MP_START, 0&, 0&)
 
    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
         
        TempPassingTime = Timer - TempOldTime
       
    Loop Until AlcorMPMessage = WM_CAM_MP_PASS _
          Or AlcorMPMessage = WM_CAM_MP_UNKNOW_FAIL _
          Or AlcorMPMessage = WM_CAM_CSET_PASS _
          Or AlcorMPMessage = WM_CAM_FWCHECK_FAIL _
          Or TempPassingTime > 10

    If AlcorMPMessage = WM_CAM_MP_PASS Then
        Image_Test = 1
    Else
        Image_Test = 0
    End If
    
    If AlcorMPMessage = WM_CAM_FWCHECK_FAIL Then
        FW_Fail_Flag = True
    End If
    
End Function

Public Sub GPIO_Setting(wPar As Long, lPar As Long)
Dim rt2

    winHwnd = FindWindow(vbNullString, "VideoCap")
    rt2 = PostMessage(winHwnd, WM_CAM_GPIO_SETTING, wPar, ByVal lPar)
 
End Sub

Public Function GPIO_Read(wPar As Long, lPar As Long) As Long
Dim rt2
Dim TempPassingTime As Long
Dim TempOldTime As Long
Dim mMsg As MSG
    
    GPIO_Read = 0
    winHwnd = FindWindow(vbNullString, "VideoCap")
    TempOldTime = Timer
    rt2 = PostMessage(winHwnd, WM_CAM_GPIO_READ, wPar, ByVal lPar)
    
'    Do
'        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
'            AlcorMPMessage = mMsg.message
'            TranslateMessage mMsg
'            DispatchMessage mMsg
'        End If
'
'        TempPassingTime = Timer - TempOldTime
'
'    Loop Until AlcorMPMessage = WM_CAM_GPIO_READ _
'          Or TempPassingTime > 5
 
    Do
        WaitMessage
        If PeekMessage(mMsg, 0, WM_CAM_GPIO_READ, WM_CAM_GPIO_READ, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If

        TempPassingTime = Timer - TempOldTime

    Loop Until AlcorMPMessage = WM_CAM_GPIO_READ _
          Or TempPassingTime > 5
 
    If mMsg.wParam <> wPar Then
        GPIO_Read = 255
    ElseIf mMsg.lParam = 0 Then
        GPIO_Read = 0
    Else
        GPIO_Read = mMsg.lParam
    End If
    
End Function

Public Sub StartROMTest_Click_AU3821()

Dim rt2
    winHwnd = FindWindow(vbNullString, "VideoCap")
    rt2 = PostMessage(winHwnd, WM_CAM_SC_START, 0&, 0&)

End Sub

Public Sub AU3821B54CFF20TestSub()

Dim OldTimer
Dim PassTime
Dim rt2
Dim LightOn
Dim mMsg As MSG
Dim LedCount As Byte

    'add unload driver function
     If PCI7248InitFinish = 0 Then
           PCI7248Exist
     End If

 AlcorMPMessage = 0
 
 cardresult = DO_WritePort(card, Channel_P1A, &H7F) 'Open ENA Power 0111_1111
 MsecDelay (2#)
 
   MPTester.TestResultLab = ""
'===============================================================
' Fail location initial
'===============================================================
 
 'AU7510 do not have filter driver
 
 
     
                     


NewChipFlag = 0
If OldChipName <> ChipName Then

' reset program

    winHwnd = FindWindow(vbNullString, "VideoCap")
 
    If winHwnd <> 0 Then
        
        Do
            rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
            Call MsecDelay(0.5)
            winHwnd = FindWindow(vbNullString, "VideoCap")
        Loop While winHwnd <> 0
    
    End If

            FileCopy App.Path & "\CamTest\" & ChipName & "\allow.sys", "C:\WINDOWS\system32\drivers\allow.sys"
            Shell App.Path & "\CamTest\" & ChipName & "\ALInstFtr -i allow 3823"
            Call MsecDelay(1#)
           ' FileCopy App.Path & "\AlcorMP_698x_PD\RAM\" & chipname & "\RAM.Bin", App.Path & "\AlcorMP_698x_PD\RAM.Bin"
           ' FileCopy App.Path & "\AlcorMP_698x_PD\INI\" & chipname & "\AlcorMP.ini", App.Path & "\AlcorMP_698x_PD\AlcorMP.ini"
            NewChipFlag = 1 ' force MP

End If
          
OldChipName = ChipName
 

 
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail


If NewChipFlag = 1 Or FindWindow(vbNullString, "VideoCap") = 0 Then
    
    MPTester.Print "wait for VideoCap Ready"
    Call LoadMP_Click_AU3830
       
End If
 
    OldTimer = Timer
    AlcorMPMessage = 0
        
    Do
        ' DoEvents
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
        
        PassTime = Timer - OldTimer
    
    Loop Until AlcorMPMessage = WM_CAM_MP_READY Or PassTime > 30 _
          Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
              
        
    MPTester.Print "Ready Time="; PassTime
        
       
    If PassTime > 15 Then    'usb issue so when time out , we let restart PC
    
    'restart PC
        MPTester.TestResultLab = "Bin3:VideoCap Ready Fail "
        TestResult = "Bin3"
        MPTester.Print "VideoCap Ready Fail"
   
        Exit Sub
   
    End If
         
         
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        
        Call StartRWTest_Click_AU3821
         
        Do
             If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
             
            PassTime = Timer - OldTimer
           
        Loop Until AlcorMPMessage = WM_CAM_MP_PASS _
              Or AlcorMPMessage = WM_CAM_MP_UNKNOW_FAIL _
              Or AlcorMPMessage = WM_CAM_MP_GPIO_FAIL _
              Or PassTime > 20 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
    
        MPTester.Print "RW work Time="; PassTime
        MPTester.MPText.Text = Hex(AlcorMPMessage)
        
        
        '===========================================================
        '  RW Time Out Fail
        '===========================================================
        
        If PassTime > 20 Then
            TestResult = "Bin3"
            MPTester.TestResultLab = "Bin3:RW Time Out Fail"
   
       
            Exit Sub
        End If
        
        ''debug.print AlcorMPMessage
        cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
        MsecDelay (0.2)
               
        Select Case AlcorMPMessage
            
            Case WM_CAM_MP_UNKNOW_FAIL
                TestResult = "Bin2"
                MPTester.TestResultLab = "Bin2:UnKnow Fail"
                ContFail = ContFail + 1
        
            Case WM_CAM_MP_GPIO_FAIL
                TestResult = "Bin3"
                MPTester.TestResultLab = "Bin3:GPIO Error "
                ContFail = ContFail + 1
    
            Case WM_CAM_MP_PASS
        
                MPTester.TestResultLab = "PASS "
                TestResult = "PASS"
                ContFail = 0
            Case Else
             
                TestResult = "Bin2"
                MPTester.TestResultLab = "Bin2:Undefine Fail"
                ContFail = ContFail + 1
        
        End Select
         
                            
End Sub

 Public Sub AU3821A66XNF20TestSub()
'add unload driver function
 If PCI7248InitFinish = 0 Then
       PCI7248Exist
 End If
 
 Dim OldTimer
 Dim PassTime
 Dim rt2
 Dim LightOn
 Dim mMsg As MSG
 Dim LedCount As Byte
 Dim DevEum As Integer
 Dim ParaString As String
 Dim EPMResult As Long
 Dim EPMString As String
 Dim EPMExe As String
 Dim EPMPath As String
 
 AlcorMPMessage = 0
 
 
 If OldChipName <> ChipName Then
 
    winHwnd = FindWindow(vbNullString, "VideoCap")
 
    If winHwnd <> 0 Then
        
        Do
            rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
            Call MsecDelay(0.5)
            winHwnd = FindWindow(vbNullString, "VideoCap")
        Loop While winHwnd <> 0
    
    End If
    
    
    FileCopy App.Path & "\CamTest\" & ChipName & "\allow.sys", "C:\WINDOWS\system32\drivers\allow.sys"
    FileCopy App.Path & "\CamTest\" & ChipName & "\ALInstFtr.exe", App.Path & "\ALInstFtr.exe"
    'Shell App.Path & "\CamTest\" & chipname & "\ALInstFtr -i allow " & ParaString
    Call MsecDelay(1#)
    
    NewChipFlag = 1 ' force MP

    cardresult = DO_WritePort(card, Channel_P1A, &HFA) 'Open ENA & EEPROM_WP 1111_1010
    Call MsecDelay(0.2)
    DevEum = WaitDevOn("vid")
    Call MsecDelay(0.4)
    
    
    If Dir(App.Path & "\Camtest\" & ChipName & "\UpdatePASS") = "UpdatePASS" Then
        Kill (App.Path & "\Camtest\" & ChipName & "\UpdatePASS")
        Call MsecDelay(0.2)
    End If
    
    If Dir(App.Path & "\Camtest\" & ChipName & "\UpdateFAIL") = "UpdateFAIL" Then
        Kill (App.Path & "\Camtest\" & ChipName & "\UpdateFAIL")
        Call MsecDelay(0.2)
    End If
    
    EPMString = App.Path & "\CamTest\" & ChipName & "\UpdateEEPROM.BAT"
    EPMExe = App.Path & "\CamTest\" & ChipName & "\AlcorMPTool_Lite_v3.0.17.1.exe"
    EPMPath = App.Path & "\CamTest\" & ChipName
    
    EPMResult = Shell(EPMString & " " & EPMExe & " " & EPMPath, vbNormalFocus)
    
    Call MsecDelay(8#)
    
    cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'Close power
    Call MsecDelay(5#)
    
    If Dir(App.Path & "\Camtest\" & ChipName & "\UpdateFAIL") = "UpdateFAIL" Then
        Kill (App.Path & "\Camtest\" & ChipName & "\UpdateFAIL")
        Call MsecDelay(0.2)
        MPTester.Print "Update EEPROM Fail!"
        cardresult = DO_WritePort(card, Channel_P1A, &HFE)
        GoTo AU3821TestResult
    ElseIf Dir(App.Path & "\Camtest\" & ChipName & "\UpdatePASS") = "UpdatePASS" Then
        Kill (App.Path & "\Camtest\" & ChipName & "\UpdatePASS")
        Call MsecDelay(0.2)
        MPTester.Print "Update EEPROM Success!"
    Else
        MPTester.Print "Update EEPROM Unknow!"
        cardresult = DO_WritePort(card, Channel_P1A, &HFE)
        GoTo AU3821TestResult
    End If
    
    cardresult = DO_WritePort(card, Channel_P1A, &HFC) 'Open ENA Power 1111_1110 check LED
    Call MsecDelay(0.2)
    DevEum = WaitDevOn("vid")
    EPMResult = Shell(App.Path & "\CamTest\" & ChipName & "\UpdateFilter.BAT", vbNormalFocus)
    Call MsecDelay(1#)
    cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Open ENA Power 1111_1110 check LED
    Call MsecDelay(0.2)
    MPTester.Print "Update Driver OK!"
    
 End If
 
 cardresult = DO_WritePort(card, Channel_P1A, &HFE) 'Open ENA Power 1111_1110 check LED
 Call MsecDelay(0.1)
 DevEum = WaitDevOn("vid")
 
 If Not DevEum Then
    GoTo AU3821TestResult
 End If
 
 cardresult = DO_ReadPort(card, Channel_P1B, LightOn)   'Get LED value
 Call MsecDelay(0.1)
 
 cardresult = DO_WritePort(card, Channel_P1A, &HC) 'Open ENA Power 1111_1110
 Call MsecDelay(0.1)
 
   MPTester.TestResultLab = ""
'===============================================================
' Fail location initial
'===============================================================
 
NewChipFlag = 0

          
OldChipName = ChipName
 
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail


If NewChipFlag = 1 Or FindWindow(vbNullString, "VideoCap") = 0 Then
    
    MPTester.Print "wait for VideoCap Ready"
    Call LoadMP_Click_AU3821
       
End If
 
    OldTimer = Timer
    AlcorMPMessage = 0
        
    Do
        ' DoEvents
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
        
        PassTime = Timer - OldTimer
    
    Loop Until AlcorMPMessage = WM_CAM_MP_READY Or PassTime > 30 _
          Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
              
        
    MPTester.Print "Ready Time="; PassTime
        
       
    If PassTime > 15 Then    'usb issue so when time out , we let restart PC
    
    'restart PC
        MPTester.TestResultLab = "Bin3:VideoCap Ready Fail "
        TestResult = "Bin3"
        MPTester.Print "VideoCap Ready Fail"
   
        Exit Sub
   
    End If
         
         
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        
        Call StartRWTest_Click_AU3821
         
        Do
             If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
             
            PassTime = Timer - OldTimer
           
        Loop Until AlcorMPMessage = WM_CAM_MP_PASS _
              Or AlcorMPMessage = WM_CAM_MP_UNKNOW_FAIL _
              Or AlcorMPMessage = WM_CAM_MP_GPIO_FAIL _
              Or PassTime > 20 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
    
        MPTester.Print "RW work Time="; PassTime
        MPTester.MPText.Text = Hex(AlcorMPMessage)
        
        
        '===========================================================
        '  RW Time Out Fail
        '===========================================================
        
        If PassTime > 20 Then
            TestResult = "Bin3"
            MPTester.TestResultLab = "Bin3:RW Time Out Fail"
            Exit Sub
        End If
        
        ''debug.print AlcorMPMessage
        
               
AU3821TestResult:
        cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
        MsecDelay (0.2)
        
        Select Case AlcorMPMessage
            
            Case WM_CAM_MP_UNKNOW_FAIL, 0
                TestResult = "Bin2"
                MPTester.TestResultLab = "Bin2:UnKnow Fail"
                ContFail = ContFail + 1
        
            Case WM_CAM_MP_GPIO_FAIL
                TestResult = "Bin3"
                MPTester.TestResultLab = "Bin3:GPIO Error "
                ContFail = ContFail + 1
    
            Case WM_CAM_MP_PASS
                
                If LightOn = &HFC Then
                    MPTester.TestResultLab = "PASS "
                    TestResult = "PASS"
                    ContFail = 0
                Else
                    MPTester.TestResultLab = "Bin4:V18 Fail"
                    TestResult = "Bin4"
                    ContFail = ContFail + 1
                End If
                
                
            Case Else
             
                TestResult = "Bin2"
                MPTester.TestResultLab = "Bin2:Undefine Fail"
                ContFail = ContFail + 1
        
        End Select
         
                            
End Sub

Public Sub AU3821A66FNF21TestSub()
'add unload driver function
 If PCI7248InitFinish = 0 Then
       Call PCI7248Exist
 End If
 
 Dim i As Integer
 Dim OldTimer As Long
 Dim PassTime As Long
 Dim rt2 As Long
 Dim LDOValue As Long
 Dim mMsg As MSG
 Dim TempResult As Byte
 Dim GPIO_Value As Long
 Dim TempCount As Integer
 
 TestResult = ""
 TempResult = 0
 AlcorMPMessage = 0
 
cardresult = DO_WritePort(card, Channel_P1A, &HFE) 'Open ENA Power 1111_1110

Call MsecDelay(0.2)
If Not WaitDevOn("vid_058f") Then
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2:Vid/Pid UnKnow Fail"
    cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
    Exit Sub
End If
 
Call MsecDelay(0.2)
 
MPTester.TestResultLab = ""
'===============================================================
' Fail location initial
'===============================================================

If OldChipName <> ChipName Then
    
    ChDir App.Path & "\CamTest\AU3821A66FNF21\"
    
    If Dir("C:\WINDOWS\system32\drivers\allow.sys") = "allow.sys" Then
        Kill ("C:\WINDOWS\system32\drivers\allow.sys")
    End If
    
    Call CloseVedioCap
    OldChipName = ChipName
End If


If FindWindow(vbNullString, "VideoCap") = 0 Then
    
    MPTester.Print "wait for VideoCap Ready"
    
    OldTimer = Timer
    
    If LoadVedioCap_AU3821 Then
        MPTester.Print "Ready Time="; Timer - OldTimer
    Else
        MPTester.TestResultLab = "Bin2:VideoCap Ready Fail "
        TestResult = "Bin2"
        MPTester.Print "VideoCap Ready Fail"
        cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
        Exit Sub
   End If
End If
         

'TempResult
'Bit 1,2  : GPIO_Setting        11: PASS
'Bit 3    : Image pattern        1: PASS

'=====================================
'   GPIO Setting
'=====================================
MPTester.Print "Begin GPIO Setting Test........"
TempCount = 0

Do
    Call GPIO_Setting(&H20, &H0)
    Call MsecDelay(0.2)
    cardresult = DO_ReadPort(card, Channel_P1B, GPIO_Value)
    Call MsecDelay(0.02)

    If (CByte(GPIO_Value) And (&H1)) = &H0 Then
        TempResult = TempResult + 1
        Exit Do
    End If
    
    TempCount = TempCount + 1
    Call MsecDelay(0.02)
    
Loop Until (TempCount > 10)


TempCount = 0

Do
    Call GPIO_Setting(&H20, &HFF)
    Call MsecDelay(0.2)
    cardresult = DO_ReadPort(card, Channel_P1B, GPIO_Value)
    Call MsecDelay(0.02)
    
     If (CByte(GPIO_Value) And (&H1)) = &H1 Then
        TempResult = TempResult + 2
        Exit Do
    End If
    
    TempCount = TempCount + 1
    Call MsecDelay(0.02)
Loop Until (TempCount > 10)


If (TempResult And &H3) = 3 Then
    MPTester.Print "GPIO Setting: PASS"
Else
    GoTo TestEnd
End If

'=====================================
'   Image
'=====================================
Call MsecDelay(0.2)
MPTester.Print "Begin Image Test........"

TempResult = TempResult + (Image_Test * &H4)

If (TempResult And &H4) = &H4 Then
    MPTester.Print "Image Test: PASS"
End If


TestEnd:

cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
WaitDevOFF ("058f")
Call MsecDelay(0.2)

If (TempResult And &H3) <> &H3 Then
    TestResult = "Bin3"
    MPTester.TestResultLab = "Bin3: GPIO Setting Fail"
ElseIf ((TempResult And &H4) <> &H4) Then
    TestResult = "Bin4"
    MPTester.TestResultLab = "Bin4: Image Fail"
ElseIf (TempResult = &H7) Then
    TestResult = "PASS"
    MPTester.TestResultLab = "Bin1: PASS"
Else
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2: Undefine Fail"
End If
                            
If (TestResult <> "PASS") And FailCloseAP Then
    Call CloseVedioCap
End If
                            
End Sub

Public Sub AU3825A61FTTestSub()
'add unload driver function
 If PCI7248InitFinish = 0 Then
       PCI7248Exist
 End If
 
 Dim OldTimer
 Dim PassTime
 Dim rt2
 Dim LightOn
 Dim mMsg As MSG
 Dim LedCount As Byte
 Dim TmpStr As String
 Dim EEPROM_Res As Long
 
 TestResult = ""
 
 AlcorMPMessage = 0
 
cardresult = DO_WritePort(card, Channel_P1A, &HFE) 'Open ENA Power 1111_1110
 
Call MsecDelay(0.6)
If Not WaitDevOn("vid_058f") Then
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2:Vid/Pid UnKnow Fail"
    cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
    ContFail = ContFail + 1
    Exit Sub
End If
 
Call MsecDelay(0.2)

   MPTester.TestResultLab = ""
'===============================================================
' Fail location initial
'===============================================================



If OldChipName <> ChipName Then
    'NewChipFlag = 0

' reset program

    winHwnd = FindWindow(vbNullString, "VideoCap")
 
    If winHwnd <> 0 Then
        
        Do
            rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
            Call MsecDelay(0.5)
            winHwnd = FindWindow(vbNullString, "VideoCap")
        Loop While winHwnd <> 0
    
    End If
    
    ChDir App.Path & "\CamTest\" & ChipName & "\"
    
    FileCopy App.Path & "\CamTest\" & ChipName & "\allow.sys", "C:\WINDOWS\system32\drivers\allow.sys"
    Shell App.Path & "\CamTest\" & ChipName & "\ALInstFtr -d allow 3823"
    Call MsecDelay(0.5)
    Shell App.Path & "\CamTest\" & ChipName & "\ALInstFtr -i allow 3823"
    Call MsecDelay(0.5)
    
    If Dir(App.Path & "\CamTest\" & ChipName & "\status.txt") = "status.txt" Then
        Kill (App.Path & "\CamTest\" & ChipName & "\status.txt")
        Call MsecDelay(0.3)
    End If
            
    'MPTester.Print "VerifyFW ..."
    '
    'If WaitDevOn("058f") Then
    '
    '    Call Load_VerifyFW_Tool_AU3825
    '
    '    OldTimer = Timer
    '
    '    Call MsecDelay(1#)
    '    winHwnd = FindWindow(vbNullString, "AlcorMPTool v3.2.4")
    '
    '    If winHwnd <> 0 Then
    '        Do
    '            Call MsecDelay(0.5)
    '            PassTime = Timer - OldTimer
    '            winHwnd = FindWindow(vbNullString, "AlcorMPTool v3.2.4")
    '        Loop While (winHwnd <> 0) Or (PassTime > 10)
    '
    '    End If
    '
    '    Call MsecDelay(0.5)
    '
    '    If Dir(App.Path & "\CamTest\" & ChipName & "\status.txt") = "status.txt" Then
    '        Open App.Path & "\CamTest\" & ChipName & "\status.txt" For Input As #5
    '        Input #5, TmpStr
    '        Close #5
    '
    '        If Trim(TmpStr) <> "FW Check Pass" Then
    '            MsgBox ("Please Check FW Version (EEPROM) !!")
    '            End
    '        End If
    '    Else
    '        TestResult = "Bin2"
    '        MPTester.TestResultLab = "Bin2:UnKnow Fail"
    '        ContFail = ContFail + 1
    '        Exit Sub
    '    End If
    '
    '    OldChipName = ChipName
    'Else
    '    TestResult = "Bin2"
    '    MPTester.TestResultLab = "Bin2:UnKnow Fail"
    '    ContFail = ContFail + 1
    '    Exit Sub
    'End If
    OldChipName = ChipName
End If
          
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail


If FindWindow(vbNullString, "VideoCap") = 0 Then
    
    MPTester.Print "wait for VideoCap Ready"
    Call LoadMP_Click_AU3830
    
    OldTimer = Timer
    AlcorMPMessage = 0
        
    Do
        ' DoEvents
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
        
        PassTime = Timer - OldTimer
    
    Loop Until AlcorMPMessage = WM_CAM_MP_READY Or PassTime > 15 _
          Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
              
        
    MPTester.Print "Ready Time="; PassTime
        
    If PassTime > 15 Then    'usb issue so when time out , we let restart PC
    
    'restart PC
        MPTester.TestResultLab = "Bin3:VideoCap Ready Fail "
        TestResult = "Bin3"
        MPTester.Print "VideoCap Ready Fail"
        cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
        Exit Sub
   
    End If
    
End If
         
OldTimer = Timer
AlcorMPMessage = 0
MPTester.Print "RW Tester begin test........"

Call StartRWTest_Click_AU3821
 
Do
    If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
        AlcorMPMessage = mMsg.message
        TranslateMessage mMsg
        DispatchMessage mMsg
    End If
     
    PassTime = Timer - OldTimer
   
Loop Until AlcorMPMessage = WM_CAM_MP_PASS _
      Or AlcorMPMessage = WM_CAM_MP_UNKNOW_FAIL _
      Or AlcorMPMessage = WM_CAM_MP_GPIO_FAIL _
      Or PassTime > 5 _
      Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY

MPTester.Print "RW work Time="; PassTime
MPTester.MPText.Text = Hex(AlcorMPMessage)

If PassTime > 5 Then
    TestResult = "Bin3"
    MPTester.TestResultLab = "Bin3:RW Time Out Fail"
    cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
    Exit Sub
End If


'=============== Check ROM Code ================

If (AlcorMPMessage = WM_CAM_MP_PASS) And (TestResult = "") Then
    Do                                                          '0000 0000
        cardresult = DO_ReadPort(card, Channel_P1B, LightOn)    'Get LED value
        Call MsecDelay(0.1)
        LedCount = LedCount + 1
        ''debug.print Hex(LightOn)
    Loop Until ((LedCount > 5) Or (LightOn = &H0))
    
    cardresult = DO_WritePort(card, Channel_P1A, &HFC) 'Close ENA Power 1111_1100
    Call MsecDelay(0.2)
    
    OldTimer = Timer
    AlcorMPMessage = 0
    MPTester.Print "Begin ROM Check........"
    
    Call StartROMTest_Click_AU3821
     
    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
         
        PassTime = Timer - OldTimer
       
    Loop Until AlcorMPMessage = WM_CAM_SC_PASS _
          Or AlcorMPMessage = WM_CAM_SC_FAIL _
          Or PassTime > 5 _
          Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY

    MPTester.Print "RW work Time="; PassTime
    MPTester.MPText.Text = Hex(AlcorMPMessage)
    
    If PassTime > 5 Then
        TestResult = "Bin3"
        MPTester.TestResultLab = "Bin3: Check ROM Time Out Fail"
        cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
        Exit Sub
    End If
     
End If

''debug.print AlcorMPMessage
cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
Call MsecDelay(0.2)
       
Select Case AlcorMPMessage
    
    Case WM_CAM_MP_UNKNOW_FAIL
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2:UnKnow Fail"
        ContFail = ContFail + 1

    Case WM_CAM_MP_GPIO_FAIL
        TestResult = "Bin3"
        MPTester.TestResultLab = "Bin3:GPIO Error "
        ContFail = ContFail + 1
    
    Case WM_CAM_SC_FAIL
        TestResult = "Bin4"
        MPTester.TestResultLab = "Bin4:ROM Fail"
        ContFail = ContFail + 1
        
    Case WM_CAM_SC_PASS
        
        If LightOn = &H0 Then
            TestResult = "PASS"
            MPTester.TestResultLab = "PASS"
            ContFail = 0
        ElseIf (LightOn = &H40) Or (LightOn = &H20) Then
            TestResult = "Bin5"
            MPTester.TestResultLab = "Bin5:Secondary PASS"
            ContFail = 0
        Else
            MPTester.TestResultLab = "Bin3: LDO Fail"
            TestResult = "Bin3"
            ContFail = ContFail + 1
        End If
            
    Case Else
     
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2:Undefine Fail"
        ContFail = ContFail + 1

End Select
                            
End Sub

Public Sub AU3825A61BFQ2ETestSub()

'2012/8/13 EQC fail lot sorting program

If PCI7248InitFinish = 0 Then
      Call PCI7248Exist
End If

If Not SetP1CInput_Flag Then
   result = DIO_PortConfig(card, Channel_P1C, INPUT_PORT)
   If result <> 0 Then
       MsgBox " config PCI_P1C as input card fail"
       End
   End If
   SetP1CInput_Flag = True
End If
 
Dim i As Integer
Dim OldTimer As Long
Dim PassTime As Long
Dim rt2 As Long
Dim LDOValue As Long
Dim mMsg As MSG
Dim LDORetry As Byte
Dim TmpStr As String
Dim TempResult As Byte
Dim GPIO_Value As Long
Dim TempCount As Byte
Dim SRAMPASSCount As Byte
Dim V18FailCount As Integer
Dim SecondaryCount As Integer
Dim LDOPASSCount As Integer
ReDim LDOVal(1 To 50) As Long
 
TestResult = ""
TempResult = 0
LDORetry = 0
AlcorMPMessage = 0
FW_Fail_Flag = False
 
cardresult = DO_WritePort(card, Channel_P1A, &HFA) 'Open ENA Power 1111_1010 (Bit3 using External clock)

Call MsecDelay(0.2)
If Not WaitDevOn("vid_058f") Then
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2:Vid/Pid UnKnow Fail"
    cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
    Call CloseVedioCap
    Exit Sub
End If
Call MsecDelay(0.8)
MPTester.TestResultLab = ""
'===============================================================
' Fail location initial
'===============================================================

If OldChipName <> ChipName Then
    
    ChDir App.Path & "\CamTest\AU3825A61FTTest_40QFN\"
    
    If Dir("C:\WINDOWS\system32\drivers\allow.sys") = "allow.sys" Then
        Kill ("C:\WINDOWS\system32\drivers\allow.sys")
    End If
    
    FWFail_Counter = 0
    
    Call CloseVedioCap
    OldChipName = ChipName

    Call Load_VerifyFW_Tool_AU3825_40QFN
    
    KillProcess ("VerifyFW_v3.2.4.exe")
End If

If FindWindow(vbNullString, "VideoCap") = 0 Then
    
    MPTester.Print "wait for VideoCap Ready"
    
    OldTimer = Timer
    
    If LoadVedioCap_AU3825_40QFN Then
        MPTester.Print "Ready Time="; Timer - OldTimer
    Else
        MPTester.TestResultLab = "Bin2:VideoCap Ready Fail "
        TestResult = "Bin2"
        MPTester.Print "VideoCap Ready Fail"
        cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
        Call CloseVedioCap
        Exit Sub
   End If
End If
         

'TempResult
'Bit 1,2  : GPIO_Setting        11: PASS
'Bit 3    : Image pattern        1: PASS
'Bit 4,5  : LDO                 11: PASS, 01:Condition PASS
'Bit 6    : SRAM                 1: PASS
'Bit 7,8    : CSET Value        01: PASS

'======================================
'   Set LV & SRAM Test
'======================================
cardresult = DO_WritePort(card, Channel_P1A, &HF8) 'Select External、Enable External 1.62V、Open ENA Power XTAL 1111_1000
Call MsecDelay(0.8)

For SRAMPASSCount = 1 To 7
    If SRAM_Test = 0 Then
        MPTester.Print "SRAM Test " & " Cycle " & SRAMPASSCount & ": Fail"
        Exit For
    Else
        MPTester.Print "Cycle " & SRAMPASSCount & ": PASS"
    End If
    
    If SRAMPASSCount = 7 Then
        TempResult = TempResult + (SRAM_Test * &H20)

        If (TempResult And &H20) = &H20 Then
            MPTester.Print "SRAM Test: PASS"
        End If
    End If
Next

If (TempResult And &H20) <> &H20 Then
    GoTo TestEnd
End If

cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Power OFF
WaitDevOFF ("vid_058f")
Call MsecDelay(0.2)
cardresult = DO_WritePort(card, Channel_P1A, &HFE) 'Power ON (Ena、Internal RC)
Call MsecDelay(0.2)


If Not WaitDevOn("vid_058f") Then
    TestResult = "Bin4"
    MPTester.TestResultLab = "Bin4:Internal UnKnow Fail"
    cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
    Exit Sub
End If

TestEnd:

cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111

WaitDevOFF ("058f")
WaitDevOFF ("058f")
Call MsecDelay(0.2)

If (TempResult And &H20) <> &H20 Then
    TestResult = "Bin3"
    MPTester.TestResultLab = "Bin3: SRAM Fail"

ElseIf (TempResult And &H20) = &H20 Then
    TestResult = "PASS"
    MPTester.TestResultLab = "Bin1: PASS"
Else
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2: Undefine Fail"
End If
                            
If TestResult = "Bin2" And FailCloseAP Then
    Call CloseVedioCap
End If
                            
End Sub

Public Sub AU3825A61BFF2FTestSub()
'add unload driver function

If PCI7248InitFinish = 0 Then
      Call PCI7248Exist
End If

If Not SetP1CInput_Flag Then
   result = DIO_PortConfig(card, Channel_P1C, INPUT_PORT)
   If result <> 0 Then
       MsgBox " config PCI_P1C as input card fail"
       End
   End If
   SetP1CInput_Flag = True
End If
 
Dim i As Integer
Dim OldTimer As Long
Dim PassTime As Long
Dim rt2 As Long
Dim LDOValue As Long
Dim mMsg As MSG
Dim LDORetry As Byte
Dim TmpStr As String
Dim TempResult As Byte
Dim GPIO_Value As Long
Dim TempCount As Byte
Dim SRAMPASSCount As Byte
Dim V18FailCount As Integer
Dim SecondaryCount As Integer
Dim LDOPASSCount As Integer
ReDim LDOVal(1 To 50) As Long
 
TestResult = ""
TempResult = 0
LDORetry = 0
AlcorMPMessage = 0
FW_Fail_Flag = False
 
cardresult = DO_WritePort(card, Channel_P1A, &HFA) 'Open ENA Power 1111_1010 (Bit3 using External clock)

Call MsecDelay(0.2)
If Not WaitDevOn("vid_058f") Then
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2:Vid/Pid UnKnow Fail"
    cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
    Call CloseVedioCap
    Exit Sub
End If
Call MsecDelay(0.8)
MPTester.TestResultLab = ""
'===============================================================
' Fail location initial
'===============================================================

If OldChipName <> ChipName Then
    
    ChDir App.Path & "\CamTest\AU3825A61FTTest_40QFN\"
    
    If Dir("C:\WINDOWS\system32\drivers\allow.sys") = "allow.sys" Then
        Kill ("C:\WINDOWS\system32\drivers\allow.sys")
    End If
    
    FWFail_Counter = 0
    
    Call CloseVedioCap
    OldChipName = ChipName

    Call Load_VerifyFW_Tool_AU3825_40QFN
    
    KillProcess ("VerifyFW_v3.2.4.exe")
End If

If FindWindow(vbNullString, "VideoCap") = 0 Then
    
    MPTester.Print "wait for VideoCap Ready"
    
    OldTimer = Timer
    
    If LoadVedioCap_AU3825_40QFN Then
        MPTester.Print "Ready Time="; Timer - OldTimer
    Else
        MPTester.TestResultLab = "Bin2:VideoCap Ready Fail "
        TestResult = "Bin2"
        MPTester.Print "VideoCap Ready Fail"
        cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
        Call CloseVedioCap
        Exit Sub
   End If
End If
         

'TempResult
'Bit 1,2  : GPIO_Setting        11: PASS
'Bit 3    : Image pattern        1: PASS
'Bit 4,5  : LDO                 11: PASS, 01:Condition PASS
'Bit 6    : SRAM                 1: PASS
'Bit 7,8    : CSET Value        01: PASS
                 
'=====================================
'   Set ST1 Level3
'=====================================
'Call GPIO_Setting(&H564, &H64)
'Call MsecDelay(0.02)
'Call GPIO_Setting(&H52A, &HB)
'Call MsecDelay(0.02)
'Call GPIO_Setting(&H52B, &H78)
'Call MsecDelay(0.02)


'=====================================
'   GPIO Setting
'=====================================
MPTester.Print "Begin GPIO Setting Test........"

TempCount = 0

Do
    Call GPIO_Setting(&H20, &H54)
    Call MsecDelay(0.04)
    cardresult = DO_ReadPort(card, Channel_P1C, GPIO_Value)
    Call MsecDelay(0.02)

    If GPIO_Value = &HE9 Then
        TempResult = TempResult + 1
        Exit Do
    End If
    
    TempCount = TempCount + 1
    Call MsecDelay(0.02)
    
Loop Until (TempCount > 10)


TempCount = 0
Call MsecDelay(0.02)
Do
    Call GPIO_Setting(&H20, &H29)
    Call MsecDelay(0.04)
    cardresult = DO_ReadPort(card, Channel_P1C, GPIO_Value)
    Call MsecDelay(0.02)
    
    If GPIO_Value = &HD6 Then
        TempResult = TempResult + 2
        Exit Do
    End If
    
    TempCount = TempCount + 1
    Call MsecDelay(0.02)
Loop Until (TempCount > 10)


If (TempResult And &H3) = 3 Then
    MPTester.Print "GPIO Setting: PASS"
Else
    GoTo TestEnd
End If


'=====================================
'   LDO
'=====================================
Call MsecDelay(0.2)
MPTester.Print "Begin LDO Test........"

For i = 1 To 50
    cardresult = DO_ReadPort(card, Channel_P1B, LDOVal(i))
    Call MsecDelay(0.01)
Next

For i = 1 To 50
    If (LDOVal(i) And &H3) <> 0 Then                            'VDD18 Fail
        V18FailCount = V18FailCount + 1
    ElseIf (LDOVal(i) = &H40) Or (LDOVal(i) = &H20) Then    'XSA < 2.65 or XSA > 2.95
        SecondaryCount = SecondaryCount + 1
    ElseIf LDOVal(i) = 0 Then
        LDOPASSCount = LDOPASSCount + 1
    End If
Next

'MPTester.Print "V18: " & V18FailCount
'MPTester.Print "Condition: " & SecondaryCount
'MPTester.Print "LDO PSS: " & LDOPASSCount

If V18FailCount >= 1 Then
    MPTester.Print "VDD18 Fail"
ElseIf (SecondaryCount >= 1) And (V18FailCount = 0) Then
    MPTester.Print "LDO Condition PASS"
    TempResult = TempResult + &H8
ElseIf LDOPASSCount = 50 Then
    MPTester.Print "LDO PASS"
    TempResult = TempResult + &H18
Else
    MPTester.Print "LDO Fail"
End If

'If LDOValue = 0 Then
'    MPTester.Print "LDO PASS"
'    TempResult = TempResult + &H18
'ElseIf (LDOValue = &H40) Or (LDOValue = &H20) Then      'XSA < 2.65 or XSA > 2.95
'    MPTester.Print "LDO Condition PASS"
'    TempResult = TempResult + &H8
'ElseIf (LDOValue And &H3) <> 0 Then                     'VDD18 Fail
'    MPTester.Print "Bin2: VDD18 Fail"
'    TestResult = "Bin2"
'    MPTester.TestResultLab = "Bin2: VDD18 Fail"
'    cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'Close ENA Power 1111_1111
'    TempCounter = TempCounter + 1
'    TempVDD18Counter = TempVDD18Counter + 1
'    Exit Sub
'Else
'    MPTester.Print "LDO Fail"
'End If

'Call MsecDelay(0.2)

'=====================================
'   Image
'=====================================
MPTester.Print "Begin Image Test........"
'OldTimer = Timer

TempResult = TempResult + (Image_Test * &H4)

If (TempResult And &H4) = &H4 Then
    MPTester.Print "Image Test: PASS"
Else
    GoTo TestEnd
End If

'======================================
'   Set LV & SRAM Test
'======================================

Call MsecDelay(0.3)
cardresult = DO_WritePort(card, Channel_P1A, &HF8) 'Select External、Enable External 1.62V、Open ENA Power XTAL 1111_1000
Call MsecDelay(0.8)

For SRAMPASSCount = 1 To 2
    If SRAM_Test = 0 Then
        MPTester.Print "SRAM Test " & " Cycle " & SRAMPASSCount & ": Fail"
        Exit For
    Else
        MPTester.Print "Cycle " & SRAMPASSCount & ": PASS"
    End If
    
    If SRAMPASSCount = 2 Then
        TempResult = TempResult + (SRAM_Test * &H20)

        If (TempResult And &H20) = &H20 Then
            MPTester.Print "SRAM Test: PASS"
        End If
    End If
Next

If (TempResult And &H20) <> &H20 Then
    GoTo TestEnd
End If

cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Power OFF
WaitDevOFF ("vid_058f")
Call MsecDelay(0.2)
cardresult = DO_WritePort(card, Channel_P1A, &HFE) 'Power ON (Ena、Internal RC)
Call MsecDelay(0.2)


If Not WaitDevOn("vid_058f") Then
    TestResult = "Bin4"
    MPTester.TestResultLab = "Bin4:Internal UnKnow Fail"
    cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
    Exit Sub
End If
Call MsecDelay(0.4)

'======================================
'   CSet Value Test
'======================================
MPTester.Print "Begin CSet Test........"
Call MsecDelay(0.1)
TempResult = TempResult + (CSET_Value_Test * &H40)

If winHwnd = FindWindow(vbNullString, "MPTool_lite_v3.12.620") Then
        
    If winHwnd <> 0 Then
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call CloseVedioCap
        KillProcess ("MPTool_lite_v3.12.620.exe")
        'Call LoadVedioCap_AU3825_40QFN
        'Call MsecDelay(1#)
        'TempResult = TempResult + (CSET_Value_Test * &H40)
    End If
End If

If (TempResult And &HC0) = &H40 Then
    MPTester.Print "CSet Value: PASS"
End If
Call MsecDelay(0.02)


TestEnd:

cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111

If FW_Fail_Flag Then
    FWFail_Counter = FWFail_Counter + 1
Else
    FWFail_Counter = 0
End If

If FWFail_Counter > 10 Then
    MPTester.FWFail_Label.Visible = True
Else
    MPTester.FWFail_Label.Visible = False
End If

WaitDevOFF ("058f")
WaitDevOFF ("058f")
Call MsecDelay(0.2)

If (TempResult And &H3) <> &H3 Then
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2: GPIO Setting Fail"
    'TempGPIOCounter = TempGPIOCounter + 1
ElseIf ((TempResult And &H4) <> &H4) Then
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2: Image Fail"
    'TempImageCounter = TempImageCounter + 1
    
ElseIf (TempResult And &H20) <> &H20 Then
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2: SRAM Fail"
    'TempSRAMCounter = TempSRAMCounter + 1
    
ElseIf (TempResult And &HC0) <> &H40 Then
    TestResult = "Bin4"
    MPTester.TestResultLab = "Bin4: CSet Fail"
    'TempCSETCounter = TempCSETCounter + 1

ElseIf (V18FailCount <> 0) Then
    TestResult = "Bin4"
    MPTester.TestResultLab = "Bin4: VDD18 Fail"

ElseIf (TempResult And &H18) = &H0 Then
    TestResult = "Bin3"
    MPTester.TestResultLab = "Bin3: LDO Fail"
    'TempLDOCounter = TempLDOCounter + 1
    
ElseIf (TempResult And &H18) = &H8 Then
    TestResult = "Bin5"
    MPTester.TestResultLab = "Bin5: Secondary PASS"
    'TempConditionCounter = TempConditionCounter + 1
ElseIf (TempResult = &H7F) Then
    TestResult = "PASS"
    MPTester.TestResultLab = "Bin1: PASS"
    'TempPASSCounter = TempPASSCounter + 1
Else
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2: Undefine Fail"
End If
                            
If TestResult = "Bin2" And FailCloseAP Then
    Call CloseVedioCap
End If
                            
'TempCounter = TempCounter + 1
'If TempCounter = 30 Then
'    'debug.print "PASS: " & TempPASSCounter & " ;SRAM: " & TempSRAMCounter & " ;CSET: "; TempCSETCounter _
'                ; " ;GPIO: " & TempGPIOCounter & " ;Condition: " & TempConditionCounter & " ;Image: " & TempImageCounter _
'                ; " ;LDO: " & TempLDOCounter & " ;VDD18: " & TempVDD18Counter
'
'    TempCounter = 0
'    TempSRAMCounter = 0
'    TempCSETCounter = 0
'    TempGPIOCounter = 0
'    TempConditionCounter = 0
'    TempPASSCounter = 0
'    TempImageCounter = 0
'    TempLDOCounter = 0
'    TempVDD18Counter = 0
'
'End If
                            
End Sub

Public Sub AU3826A81AFF20TestSub()
Dim i As Integer
Dim OldTimer As Long
Dim PassTime As Long
Dim rt2 As Long
Dim LDOValue As Long
Dim mMsg As MSG
Dim TmpStr As String
Dim ItemResult As Byte
Dim GPIO_Value As Long
Dim LDOPASSCount As Integer
Dim GPIOReadVal As Long
Dim GPIO15Result As Byte

Dim VDD1812Secondary As Integer
Dim VDDSASDSecondary As Integer
Dim NA As Integer
ReDim LDOVal(1 To 50) As Long

    If (InStr(1, ChipName, "AFF") <> 0 Or InStr(1, ChipName, "CFF") <> 0 Or InStr(1, ChipName, "DFF") <> 0) Then
        FileCopy App.Path & "\CamTest\AU3826A81FTTest_40QFN\3826\VideoCap.ini", App.Path & "\CamTest\AU3826A81FTTest_40QFN\VideoCap.ini"
    Else
        FileCopy App.Path & "\CamTest\AU3826A81FTTest_40QFN\3822\VideoCap.ini", App.Path & "\CamTest\AU3826A81FTTest_40QFN\VideoCap.ini"
    End If

    NA = 0
    TestResult = ""
    ItemResult = 0
    AlcorMPMessage = 0
    GPIO15Result = 0
    FW_Fail_Flag = False
    Check3826VCC5V = False

    If PCI7248InitFinish = 0 Then
        Call PCI7248Exist
    End If

    If Not SetP1CInput_Flag Then
       result = DIO_PortConfig(card, Channel_P1C, INPUT_PORT)
       If result <> 0 Then
           MsgBox " config PCI_P1C as input card fail"
           End
       End If
       SetP1CInput_Flag = True
    End If

    cardresult = DO_WritePort(card, Channel_P1A, &HDA)      ' set &HFA to check VCC5V
    If Check3826VCC5V = False Then
        Check3826VCC5V = True
        Call MsecDelay(0.2)
        cardresult = DO_ReadPort(card, Channel_P1B, LDOVal(1))
        If LDOVal(1) = 0 Then
            MsgBox "請確認5V是否有接"
            End
        End If
    End If

    cardresult = DO_WritePort(card, Channel_P1A, &H11)      'to act like open socket behavior
    Call MsecDelay(0.2)

    ' === check device exist and then close power ===
    cardresult = DO_WritePort(card, Channel_P1A, &HCA)      'Open ENA Power 1100_1010 (Bit3 using External clock)
    Call MsecDelay(0.2)
    If Not WaitDevOn("vid_058f") Then
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2:Vid/Pid UnKnow Fail"
        cardresult = DO_WritePort(card, Channel_P1A, &HDF)  'Close ENA Power 1101_1111
        Call CloseVedioCap
        Exit Sub
    End If
    Call MsecDelay(0.7)
    MPTester.TestResultLab = ""

    ' === chipname is new or changed ===
    If OldChipName <> ChipName Then
        ChDir App.Path & "\CamTest\AU3826A81FTTest_40QFN\"

        FileCopy App.Path & "\CamTest\AU3826A81FTTest_40QFN\allow.sys", "C:\WINDOWS\system32\drivers\allow.sys"

        If (InStr(1, ChipName, "AFF") <> 0 Or InStr(1, ChipName, "CFF") <> 0 Or InStr(1, ChipName, "DFF") <> 0) Then
            Shell App.Path & "\CamTest\AU3826A81FTTest_40QFN\ALInstFtr -d allow 3826"
        Else
            Shell App.Path & "\CamTest\AU3826A81FTTest_40QFN\ALInstFtr -d allow 3822"
        End If

        Call MsecDelay(0.3)

        If (InStr(1, ChipName, "AFF") <> 0 Or InStr(1, ChipName, "CFF") <> 0 Or InStr(1, ChipName, "DFF") <> 0) Then
            Shell App.Path & "\CamTest\AU3826A81FTTest_40QFN\ALInstFtr -i allow 3826"
        Else
            Shell App.Path & "\CamTest\AU3826A81FTTest_40QFN\ALInstFtr -i allow 3822"
        End If

        Call MsecDelay(0.3)

        FWFail_Counter = 0

        Call CloseVedioCap
        OldChipName = ChipName

        Call Load_VerifyFW_Tool_AU3826A81FTTest_40QFN

        Call MsecDelay(0.5)

        KillProcess ("MPTool_lite_v3.12.620.exe")
        KillProcess ("VerifyFW_v3.2.4.exe")
    End If

    If FindWindow(vbNullString, "VideoCap") = 0 Then
        MPTester.Print "wait for VideoCap Ready"

        OldTimer = Timer

        If LoadVedioCap_AU3826A81FTTest_40QFN Then
            MPTester.Print "Ready Time="; Timer - OldTimer
        Else
            MPTester.TestResultLab = "Bin2:VideoCap Ready Fail "
            TestResult = "Bin2"
            MPTester.Print "VideoCap Ready Fail"
            cardresult = DO_WritePort(card, Channel_P1A, &HDF) 'Close ENA Power 1111_1111
            Call CloseVedioCap
            Exit Sub
       End If

    End If

    ' ======================== Test result ===========================
    '                        sav/sdv   nc   gpio15   masclk  power
    '                          LDO    EEP     ENC     ENB     ENA
    '   8       7       6       5      4       3       2       1
    '   1       1       1       0      1       0       1       0    check LDO, parallel
    '   1       1       1       1      1       1       0       0    check MIPI
    ' ================================================================

    ' ======================== ItemResult ===========================
    '(2^7) VDDSASD
    '(2^6) VDD1812
    '(2^5) LDO
    '(2^4) CSET
    '
    '(2^3) MIPI
    '(2^2) Parallel
    '(2^1) GPIO15 (LC)
    '(2^0) GPIO15 (XTAL)


    '=====================================
    '   GPIO15 Setting (set to 0) Clock source is XTAL
    '=====================================
    MPTester.Print "Begin GPIO Setting Test........"

    Do
        GPIOReadVal = GPIO_Read(&H21, 0)    'Check GPIO Value is 1
        If (GPIOReadVal = 0) Then
            Exit Do
        End If
        NA = NA + 1
    Loop While (NA < 5)

    NA = 0

    If (GPIOReadVal <> 0) Then
        MPTester.Print "Set XTAL Fail"
        GPIO15Result = 0
        GoTo TestEnd
    Else
        ItemResult = ItemResult + 1
    End If

    '=====================================
    '   Check LDO
    '=====================================
    Call MsecDelay(0.2)
    MPTester.Print "Begin LDO Test........"

    Call GPIO_Setting(&H560, &H0)
    Call MsecDelay(0.1)

    For i = 1 To 50
        cardresult = DO_ReadPort(card, Channel_P1B, LDOVal(i))
        Call MsecDelay(0.01)

        If (LDOVal(i) And &H1 = &H1) Or (LDOVal(i) And &H2 = &H2) Or (LDOVal(i) And &H4 = &H4) Or (LDOVal(i) And &H8 = &H8) Then
            VDDSASDSecondary = VDDSASDSecondary + 1
        ElseIf (LDOVal(i) And &H10 = &H10) Or (LDOVal(i) And &H20 = &H20) Or (LDOVal(i) And &H40 = &H40) Or (LDOVal(i) And &H80 = &H80) Then
            VDD1812Secondary = VDD1812Secondary + 1
        Else
            LDOPASSCount = LDOPASSCount + 1
        End If
    Next

    If LDOPASSCount = 50 Then
        MPTester.Print "LDO PASS"
        ItemResult = ItemResult + &H20
    ElseIf (VDDSASDSecondary > 1) And (VDD1812Secondary = 0) Then
        MPTester.Print "VDD1812Secondary pass"
        ItemResult = ItemResult + &H40
    ElseIf (VDDSASDSecondary = 0) And (VDD1812Secondary > 1) Then
        MPTester.Print "VDDSASDSecondary pass"
        ItemResult = ItemResult + &H80
    Else
        MPTester.Print "LDO Fail"
    End If

    '=====================================
    '   Check parallel
    '=====================================
    MPTester.Print "Begin Check Parallel........"

    ItemResult = ItemResult + (Check_Parallel * 4)

    If (ItemResult And &H4) = &H4 Then
        MPTester.Print "Check parallel: PASS"
    Else
        GoTo TestEnd
    End If

    ' ===== close power and open =====
    cardresult = DO_WritePort(card, Channel_P1A, &HD1)      'Close ENA Power 1111_0001
    WaitDevOFF ("058f")
    WaitDevOFF ("058f")
    Call MsecDelay(0.5)
    ' ================================

    cardresult = DO_WritePort(card, Channel_P1A, &HD0)      'Fine tune for stable the test
    Call MsecDelay(0.2)

    '=====================================
    '   Connect MIPI Power(Bit5 =>H) & MSACLK(Bit2 =>L)
    '=====================================
    cardresult = DO_WritePort(card, Channel_P1A, &HDC)      'Open ENA Power 1111_1100 (Bit3 using Internal LC)
    WaitDevOn ("058f")
    Call MsecDelay(0.2)

    '=====================================
    '   GPIO15 Setting (set to 1) Clock source is LC
    '=====================================
    MPTester.Print "Begin GPIO Setting Test........"

    Do
        GPIOReadVal = GPIO_Read(&H21, &HFF)    'Check GPIO Value is 1
        If (GPIOReadVal = &H80) Then
            Exit Do
        End If
        NA = NA + 1
        Call MsecDelay(0.2)
    Loop While (NA < 5)

    NA = 0

    If (GPIOReadVal <> &H80) Then
        MPTester.Print "Set LC Fail"
        GPIO15Result = 0
        GoTo TestEnd
    Else
        ItemResult = ItemResult + 2
    End If

    '=====================================
    '   Check MIPI
    '=====================================
    MPTester.Print "Begin Check MIPI........"

    ItemResult = ItemResult + (Check_MIPI * &H8)

    If (ItemResult And &H8) = &H8 Then
        MPTester.Print "Check MIPI: PASS"
    Else
        GoTo TestEnd
    End If

    '======================================
    '   CSet Value Test
    '======================================
    MPTester.Print "Begin CSet Test........"
    Call MsecDelay(0.1)
    ItemResult = ItemResult + (CSET_Value_Test * &H10)

    If (ItemResult And &H10) = &H10 Then
        MPTester.Print "CSet Value: PASS"
    Else
        GoTo TestEnd
    End If

TestEnd:
    
    If InStr(1, ChipName, "22") <> 0 Then
        cardresult = DO_WritePort(card, Channel_P1A, &HD1) 'Close ENA Power 1111_0001
    Else
        cardresult = DO_WritePort(card, Channel_P1A, &HDB) 'Close ENA Power 1111_1011
    End If

    If FW_Fail_Flag Then
        FWFail_Counter = FWFail_Counter + 1
    Else
        FWFail_Counter = 0
    End If

    If FWFail_Counter > 10 Then
        MPTester.FWFail_Label.Visible = True
    Else
        MPTester.FWFail_Label.Visible = False
    End If

    WaitDevOFF ("058f")
    WaitDevOFF ("058f")
    Call MsecDelay(0.2)

    If (ItemResult And &H1) <> &H1 Then
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2: Set XTAL Fail"

    ElseIf (ItemResult And &H40) = &H40 Then
        TestResult = "Bin4"
        MPTester.TestResultLab = "Bin4: VDD1812 Condition Pass"

    ElseIf (ItemResult And &H80) = &H80 Then
        TestResult = "Bin5"
        MPTester.TestResultLab = "Bin5: VDDSASD Condition Pass"

    ElseIf (ItemResult And &H20) <> &H20 Then
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2: LDO Fail"

    ElseIf ((ItemResult And &H4) <> &H4) Then
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2: Parallel Fail"

    ElseIf (ItemResult And &H2) <> &H2 Then
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2: Set LC Fail"

    ElseIf (ItemResult And &H8) <> &H8 Then
        TestResult = "Bin3"
        MPTester.TestResultLab = "Bin3: MIPI Fail"

    ElseIf (ItemResult And &H10) <> &H10 Then
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2: CSet Fail"

    ElseIf (ItemResult = &H3F) Then
        TestResult = "PASS"
        MPTester.TestResultLab = "Bin1: PASS"

    Else
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2: Undefine Fail"
    End If

    If TestResult = "Bin2" And FailCloseAP Then
        Call CloseVedioCap
    End If

End Sub

Public Sub AU3826A81AFF23TestSub()
Dim i As Integer
Dim OldTimer As Long
Dim PassTime As Long
Dim rt2 As Long
Dim LDOValue As Long
Dim mMsg As MSG
Dim TmpStr As String
Dim ItemResult As Byte
Dim GPIO_Value As Long
Dim LDOPASSCount As Integer
Dim GPIOReadVal As Long
Dim GPIO15Result As Byte

Dim VDD1812Secondary As Integer
Dim VDDSASDSecondary As Integer
Dim NA As Integer
ReDim LDOVal(1 To 50) As Long

    If (InStr(1, ChipName, "AFF") <> 0 Or InStr(1, ChipName, "CFF") <> 0) Then
        FileCopy App.Path & "\CamTest\AU3826A81FTTest_40QFN\3826\VideoCap.ini", App.Path & "\CamTest\AU3826A81FTTest_40QFN\VideoCap.ini"
    Else
        FileCopy App.Path & "\CamTest\AU3826A81FTTest_40QFN\3822\VideoCap.ini", App.Path & "\CamTest\AU3826A81FTTest_40QFN\VideoCap.ini"
    End If

    NA = 0
    TestResult = ""
    ItemResult = 0
    AlcorMPMessage = 0
    GPIO15Result = 0
    FW_Fail_Flag = False
    Check3826VCC5V = False

    If PCI7248InitFinish = 0 Then
        Call PCI7248Exist
    End If

    If Not SetP1CInput_Flag Then
       result = DIO_PortConfig(card, Channel_P1C, INPUT_PORT)
       If result <> 0 Then
           MsgBox " config PCI_P1C as input card fail"
           End
       End If
       SetP1CInput_Flag = True
    End If

    cardresult = DO_WritePort(card, Channel_P1A, &HDA)      ' set &HFA to check VCC5V
    If Check3826VCC5V = False Then
        Check3826VCC5V = True
        Call MsecDelay(0.1)
        cardresult = DO_ReadPort(card, Channel_P1B, LDOVal(1))
        If LDOVal(1) = 0 Then
            MsgBox "請確認5V是否有接"
            End
        End If
    End If

    cardresult = DO_WritePort(card, Channel_P1A, &H11)      'to act like open socket behavior
    Call MsecDelay(0.05)

    ' === check device exist and then close power ===
    cardresult = DO_WritePort(card, Channel_P1A, &HCA)      'Open ENA Power 1100_1010 (Bit3 using External clock)
    Call MsecDelay(0.05)
    If Not WaitDevOn("vid_058f") Then
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2:Vid/Pid UnKnow Fail"
        cardresult = DO_WritePort(card, Channel_P1A, &HDF)  'Close ENA Power 1101_1111
        Call CloseVedioCap
        Exit Sub
    End If
    Call MsecDelay(0.1)
    MPTester.TestResultLab = ""

    ' === chipname is new or changed ===
    If OldChipName <> ChipName Then
        ChDir App.Path & "\CamTest\AU3826A81FTTest_40QFN\"

        FileCopy App.Path & "\CamTest\AU3826A81FTTest_40QFN\allow.sys", "C:\WINDOWS\system32\drivers\allow.sys"

        If (InStr(1, ChipName, "AFF") <> 0 Or InStr(1, ChipName, "CFF") <> 0) Then
            Shell App.Path & "\CamTest\AU3826A81FTTest_40QFN\ALInstFtr -d allow 3826"
        Else
            Shell App.Path & "\CamTest\AU3826A81FTTest_40QFN\ALInstFtr -d allow 3822"
        End If

        Call MsecDelay(0.1)

        If (InStr(1, ChipName, "AFF") <> 0 Or InStr(1, ChipName, "CFF") <> 0) Then
            Shell App.Path & "\CamTest\AU3826A81FTTest_40QFN\ALInstFtr -i allow 3826"
        Else
            Shell App.Path & "\CamTest\AU3826A81FTTest_40QFN\ALInstFtr -i allow 3822"
        End If

        Call MsecDelay(0.1)
        FWFail_Counter = 0

        Call CloseVedioCap
        OldChipName = ChipName

        Call Load_VerifyFW_Tool_AU3826A81FTTest_40QFN
        Call MsecDelay(0.2)

        KillProcess ("MPTool_lite_v3.12.620.exe")
        KillProcess ("VerifyFW_v3.2.4.exe")
    End If

    If FindWindow(vbNullString, "VideoCap") = 0 Then
        MPTester.Print "wait for VideoCap Ready"

        OldTimer = Timer

        If LoadVedioCap_AU3826A81FTTest_40QFN Then
            MPTester.Print "Ready Time="; Timer - OldTimer
        Else
            MPTester.TestResultLab = "Bin2:VideoCap Ready Fail "
            TestResult = "Bin2"
            MPTester.Print "VideoCap Ready Fail"
            cardresult = DO_WritePort(card, Channel_P1A, &HDF) 'Close ENA Power 1111_1111
            Call CloseVedioCap
            Exit Sub
       End If

    End If

    ' ======================== Test result ===========================
    '                        sav/sdv   nc   gpio15   masclk  power
    '                          LDO    EEP     ENC     ENB     ENA
    '   8       7       6       5      4       3       2       1
    '   1       1       1       0      1       0       1       0    check LDO, parallel
    '   1       1       1       1      1       1       0       0    check MIPI
    ' ================================================================

    ' ======================== ItemResult ===========================
    '(2^7) VDDSASD
    '(2^6) VDD1812
    '(2^5) LDO
    '(2^4) CSET
    '
    '(2^3) MIPI
    '(2^2) Parallel
    '(2^1) GPIO15 (LC)
    '(2^0) GPIO15 (XTAL)


    '=====================================
    '   GPIO15 Setting (set to 0) Clock source is XTAL
    '=====================================
    MPTester.Print "Begin GPIO Setting Test........"

    Do
        GPIOReadVal = GPIO_Read(&H21, 0)    'Check GPIO Value is 1
        If (GPIOReadVal = 0) Then
            Exit Do
        End If
        NA = NA + 1
    Loop While (NA < 5)

    NA = 0

    If (GPIOReadVal <> 0) Then
        MPTester.Print "Set XTAL Fail"
        GPIO15Result = 0
        GoTo TestEnd
    Else
        ItemResult = ItemResult + 1
    End If

    '=====================================
    '   Check LDO
    '=====================================
    MPTester.Print "Begin LDO Test........"

    Call GPIO_Setting(&H560, &H0)
    Call MsecDelay(0.1)

    For i = 1 To 50
        cardresult = DO_ReadPort(card, Channel_P1B, LDOVal(i))
        Call MsecDelay(0.01)

        If (LDOVal(i) And &H1 = &H1) Or (LDOVal(i) And &H2 = &H2) Or (LDOVal(i) And &H4 = &H4) Or (LDOVal(i) And &H8 = &H8) Then
            VDDSASDSecondary = VDDSASDSecondary + 1
        ElseIf (LDOVal(i) And &H10 = &H10) Or (LDOVal(i) And &H20 = &H20) Or (LDOVal(i) And &H40 = &H40) Or (LDOVal(i) And &H80 = &H80) Then
            VDD1812Secondary = VDD1812Secondary + 1
        Else
            LDOPASSCount = LDOPASSCount + 1
        End If
    Next

    If LDOPASSCount = 50 Then
        MPTester.Print "LDO PASS"
        ItemResult = ItemResult + &H20
    ElseIf (VDDSASDSecondary > 1) And (VDD1812Secondary = 0) Then
        MPTester.Print "VDD1812Secondary pass"
        ItemResult = ItemResult + &H40
    ElseIf (VDDSASDSecondary = 0) And (VDD1812Secondary > 1) Then
        MPTester.Print "VDDSASDSecondary pass"
        ItemResult = ItemResult + &H80
    Else
        MPTester.Print "LDO Fail"
    End If

    '=====================================
    '   Check parallel
    '=====================================
    MPTester.Print "Begin Check Parallel........"

    ItemResult = ItemResult + (Check_Parallel * 4)

    If (ItemResult And &H4) = &H4 Then
        MPTester.Print "Check parallel: PASS"
    Else
        GoTo TestEnd
    End If

    ' ===== close power and open =====
    cardresult = DO_WritePort(card, Channel_P1A, &HD1)      'Close ENA Power 1111_0001
    WaitDevOFF ("058f")
    WaitDevOFF ("058f")
    Call MsecDelay(0.05)
    ' ================================

    cardresult = DO_WritePort(card, Channel_P1A, &HD0)      'Fine tune for stable the test
    Call MsecDelay(0.05)

    '=====================================
    '   Connect MIPI Power(Bit5 =>H) & MSACLK(Bit2 =>L)
    '=====================================
    cardresult = DO_WritePort(card, Channel_P1A, &HDC)      'Open ENA Power 1111_1100 (Bit3 using Internal LC)
    WaitDevOn ("058f")
    Call MsecDelay(0.05)

    '=====================================
    '   GPIO15 Setting (set to 1) Clock source is LC
    '=====================================
    MPTester.Print "Begin GPIO Setting Test........"

    Do
        GPIOReadVal = GPIO_Read(&H21, &HFF)    'Check GPIO Value is 1
        If (GPIOReadVal = &H80) Then
            Exit Do
        End If
        NA = NA + 1
        Call MsecDelay(0.05)
    Loop While (NA < 5)

    NA = 0

    If (GPIOReadVal <> &H80) Then
        MPTester.Print "Set LC Fail"
        GPIO15Result = 0
        GoTo TestEnd
    Else
        ItemResult = ItemResult + 2
    End If

    '=====================================
    '   Check MIPI
    '=====================================
    MPTester.Print "Begin Check MIPI........"

    ItemResult = ItemResult + (Check_MIPI * &H8)

    If (ItemResult And &H8) = &H8 Then
        MPTester.Print "Check MIPI: PASS"
    Else
        GoTo TestEnd
    End If

    '======================================
    '   CSet Value Test
    '======================================
    MPTester.Print "Begin CSet Test........"
    Call MsecDelay(0.05)
    ItemResult = ItemResult + (CSET_Value_Test * &H10)

    If (ItemResult And &H10) = &H10 Then
        MPTester.Print "CSet Value: PASS"
    Else
        GoTo TestEnd
    End If

TestEnd:
    
    If InStr(1, ChipName, "22") <> 0 Then
        cardresult = DO_WritePort(card, Channel_P1A, &HD1) 'Close ENA Power 1111_0001
    Else
        cardresult = DO_WritePort(card, Channel_P1A, &HDB) 'Close ENA Power 1111_1011
    End If

    If FW_Fail_Flag Then
        FWFail_Counter = FWFail_Counter + 1
    Else
        FWFail_Counter = 0
    End If

    If FWFail_Counter > 10 Then
        MPTester.FWFail_Label.Visible = True
    Else
        MPTester.FWFail_Label.Visible = False
    End If

    WaitDevOFF ("058f")
    WaitDevOFF ("058f")
    Call MsecDelay(0.05)

    If (ItemResult And &H1) <> &H1 Then
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2: Set XTAL Fail"

    ElseIf (ItemResult And &H40) = &H40 Then
        TestResult = "Bin4"
        MPTester.TestResultLab = "Bin4: VDD1812 Condition Pass"

    ElseIf (ItemResult And &H80) = &H80 Then
        TestResult = "Bin5"
        MPTester.TestResultLab = "Bin5: VDDSASD Condition Pass"

    ElseIf (ItemResult And &H20) <> &H20 Then
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2: LDO Fail"

    ElseIf ((ItemResult And &H4) <> &H4) Then
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2: Parallel Fail"

    ElseIf (ItemResult And &H2) <> &H2 Then
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2: Set LC Fail"

    ElseIf (ItemResult And &H8) <> &H8 Then
        TestResult = "Bin3"
        MPTester.TestResultLab = "Bin3: MIPI Fail"

    ElseIf (ItemResult And &H10) <> &H10 Then
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2: CSet Fail"

    ElseIf (ItemResult = &H3F) Then
        TestResult = "PASS"
        MPTester.TestResultLab = "Bin1: PASS"

    Else
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2: Undefine Fail"
    End If

    If TestResult = "Bin2" And FailCloseAP Then
        Call CloseVedioCap
    End If

End Sub

'add in 20131203 for FT3

Public Sub AU3826A81AFF30TestSub()
Dim i As Integer
Dim OldTimer As Long
Dim PassTime As Long
Dim rt2 As Long
Dim LDOValue As Long
Dim mMsg As MSG
Dim TmpStr As String
Dim ItemResult As Byte
Dim GPIO_Value As Long
Dim LDOPASSCount As Integer
Dim GPIOReadVal As Long
Dim GPIO15Result As Byte

Dim VDD1812Secondary As Integer
Dim VDDSASDSecondary As Integer
Dim NA As Integer
ReDim LDOVal(1 To 50) As Long

Dim HV As Boolean
Dim HLV_Result As Integer

    HV = True
    HLV_Result = 0

    If InStr(1, ChipName, "AFF") <> 0 Then
        FileCopy App.Path & "\CamTest\AU3826A81FTTest_40QFN\3826\VideoCap.ini", App.Path & "\CamTest\AU3826A81FTTest_40QFN\VideoCap.ini"
    Else
        FileCopy App.Path & "\CamTest\AU3826A81FTTest_40QFN\3822\VideoCap.ini", App.Path & "\CamTest\AU3826A81FTTest_40QFN\VideoCap.ini"
    End If

    NA = 0
    TestResult = ""
    ItemResult = 0
    AlcorMPMessage = 0
    GPIO15Result = 0
    FW_Fail_Flag = False
    Check3826VCC5V = False
    
    If PCI7248InitFinish = 0 Then
        Call PCI7248Exist
    End If
    
    If Not SetP1CInput_Flag Then
       result = DIO_PortConfig(card, Channel_P1C, INPUT_PORT)
       If result <> 0 Then
           MsgBox " config PCI_P1C as input card fail"
           End
       End If
       SetP1CInput_Flag = True
    End If
    
HLV_Routine:
    
    LDOPASSCount = 0
    ItemResult = 0
    
    cardresult = DO_WritePort(card, Channel_P1A, &HFA)      ' set &HFA to check VCC5V

    If Check3826VCC5V = False Then
        Check3826VCC5V = True
        Call MsecDelay(0.2)
        cardresult = DO_ReadPort(card, Channel_P1B, LDOVal(1))
        If LDOVal(1) = 0 Then
            MsgBox "請確認5V是否有接"
            End
        End If
    End If
    
    cardresult = DO_WritePort(card, Channel_P1A, &H11)      'to act like open socket behavior
    Call MsecDelay(0.2)
    
    ' === check device exist and then close power ===
    cardresult = DO_WritePort(card, Channel_P1A, &HEA)      'Open ENA Power 1110_1010 (Bit3 using External clock)
    Call MsecDelay(0.2)
    
    If HV Then
        Call PowerSet2(1, "3.6", "0.5", 1, "3.6", "0.5", 1)
    Else
        Call PowerSet2(1, "3.1", "0.5", 1, "3.1", "0.5", 1)
    End If
    
    If Not WaitDevOn("vid_058f") Then
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2:Vid/Pid UnKnow Fail"
        cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'Close ENA Power 1111_1111
        Call CloseVedioCap
        Exit Sub
    End If
    Call MsecDelay(0.7)
    MPTester.TestResultLab = ""
    
    ' === chipname is new or changed ===
    If OldChipName <> ChipName Then
        ChDir App.Path & "\CamTest\AU3826A81FTTest_40QFN\"
        
        FileCopy App.Path & "\CamTest\AU3826A81FTTest_40QFN\allow.sys", "C:\WINDOWS\system32\drivers\allow.sys"
        
        If InStr(1, ChipName, "AFF") <> 0 Then
            Shell App.Path & "\CamTest\AU3826A81FTTest_40QFN\ALInstFtr -d allow 3826"
        Else
            Shell App.Path & "\CamTest\AU3826A81FTTest_40QFN\ALInstFtr -d allow 3822"
        End If
        
        Call MsecDelay(0.3)
        
        If InStr(1, ChipName, "AFF") <> 0 Then
            Shell App.Path & "\CamTest\AU3826A81FTTest_40QFN\ALInstFtr -i allow 3826"
        Else
            Shell App.Path & "\CamTest\AU3826A81FTTest_40QFN\ALInstFtr -i allow 3822"
        End If
        
        Call MsecDelay(0.3)
    
        FWFail_Counter = 0
        
        Call CloseVedioCap
        OldChipName = ChipName
        
        Call Load_VerifyFW_Tool_AU3826A81FTTest_40QFN
        
        Call MsecDelay(0.5)
        
        KillProcess ("MPTool_lite_v3.12.620.exe")
        KillProcess ("VerifyFW_v3.2.4.exe")
    End If

    If FindWindow(vbNullString, "VideoCap") = 0 Then
        MPTester.Print "wait for VideoCap Ready"
    
        OldTimer = Timer
    
        If LoadVedioCap_AU3826A81FTTest_40QFN Then
            MPTester.Print "Ready Time="; Timer - OldTimer
        Else
            MPTester.TestResultLab = "Bin2:VideoCap Ready Fail "
            TestResult = "Bin2"
            MPTester.Print "VideoCap Ready Fail"
            cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
            Call CloseVedioCap
            Exit Sub
       End If
       
    End If
    
    ' ======================== Test result ===========================
    '                        sav/sdv   nc   gpio15   masclk  power
    '                          LDO    EEP     ENC     ENB     ENA
    '   8       7       6       5      4       3       2       1
    '   1       1       1       0      1       0       1       0    check LDO, parallel
    '   1       1       1       1      1       1       0       0    check MIPI
    ' ================================================================
    
    ' ======================== ItemResult ===========================
    '(2^7) VDDSASD
    '(2^6) VDD1812
    '(2^5) LDO
    '(2^4) CSET
    '
    '(2^3) MIPI
    '(2^2) Parallel
    '(2^1) GPIO15 (LC)
    '(2^0) GPIO15 (XTAL)
    
    
    '=====================================
    '   GPIO15 Setting (set to 0) Clock source is XTAL
    '=====================================
    MPTester.Print "Begin GPIO Setting Test........"
    
    Do
        GPIOReadVal = GPIO_Read(&H21, 0)    'Check GPIO Value is 1
        If (GPIOReadVal = 0) Then
            Exit Do
        End If
        NA = NA + 1
    Loop While (NA < 5)
    
    NA = 0
    
    If (GPIOReadVal <> 0) Then
        MPTester.Print "Set XTAL Fail"
        GPIO15Result = 0
        GoTo TestEnd
    Else
        ItemResult = ItemResult + 1
    End If
    
    '=====================================
    '   Check LDO
    '=====================================
    Call MsecDelay(0.2)
    MPTester.Print "Begin LDO Test........"

    Call GPIO_Setting(&H560, &H0)
    Call MsecDelay(0.1)
    
    For i = 1 To 50
        cardresult = DO_ReadPort(card, Channel_P1B, LDOVal(i))
        Call MsecDelay(0.01)
    
        If (LDOVal(i) And &H1 = &H1) Or (LDOVal(i) And &H2 = &H2) Or (LDOVal(i) And &H4 = &H4) Or (LDOVal(i) And &H8 = &H8) Then
            VDDSASDSecondary = VDDSASDSecondary + 1
        ElseIf (LDOVal(i) And &H10 = &H10) Or (LDOVal(i) And &H20 = &H20) Or (LDOVal(i) And &H40 = &H40) Or (LDOVal(i) And &H80 = &H80) Then
            VDD1812Secondary = VDD1812Secondary + 1
        Else
            LDOPASSCount = LDOPASSCount + 1
        End If
    Next

    If LDOPASSCount = 50 Then
        MPTester.Print "LDO PASS"
        ItemResult = ItemResult + &H20
    ElseIf (VDDSASDSecondary > 1) And (VDD1812Secondary = 0) Then
        MPTester.Print "VDD1812Secondary pass"
        ItemResult = ItemResult + &H40
    ElseIf (VDDSASDSecondary = 0) And (VDD1812Secondary > 1) Then
        MPTester.Print "VDDSASDSecondary pass"
        ItemResult = ItemResult + &H80
    Else
        MPTester.Print "LDO Fail"
    End If
    
    '=====================================
    '   Check parallel
    '=====================================
    MPTester.Print "Begin Check Parallel........"
    
    ItemResult = ItemResult + (Check_Parallel * 4)
    
    If (ItemResult And &H4) = &H4 Then
        MPTester.Print "Check parallel: PASS"
    Else
        GoTo TestEnd
    End If
    
    ' ===== close power and open =====
    cardresult = DO_WritePort(card, Channel_P1A, &HF1)      'Close ENA Power 1111_0001
    
    Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
    
    WaitDevOFF ("058f")
    WaitDevOFF ("058f")
    Call MsecDelay(0.5)
    ' ================================
    
    cardresult = DO_WritePort(card, Channel_P1A, &HF0)      'Fine tune for stable the test
    Call MsecDelay(0.2)
    
    '=====================================
    '   Connect MIPI Power(Bit5 =>H) & MSACLK(Bit2 =>L)
    '=====================================
    cardresult = DO_WritePort(card, Channel_P1A, &HFC)      'Open ENA Power 1111_1100 (Bit3 using Internal LC)
    
    If HV Then
        Call PowerSet2(1, "3.6", "0.5", 1, "3.6", "0.5", 1)
    Else
        Call PowerSet2(1, "3.1", "0.5", 1, "3.1", "0.5", 1)
    End If
    
    WaitDevOn ("058f")
    Call MsecDelay(0.2)
    
    '=====================================
    '   GPIO15 Setting (set to 1) Clock source is LC
    '=====================================
    MPTester.Print "Begin GPIO Setting Test........"
    
    Do
        GPIOReadVal = GPIO_Read(&H21, &HFF)    'Check GPIO Value is 1
        If (GPIOReadVal = &H80) Then
            Exit Do
        End If
        NA = NA + 1
        Call MsecDelay(0.2)
    Loop While (NA < 5)
    
    NA = 0
    
    If (GPIOReadVal <> &H80) Then
        MPTester.Print "Set LC Fail"
        GPIO15Result = 0
        GoTo TestEnd
    Else
        ItemResult = ItemResult + 2
    End If
    
    '=====================================
    '   Check MIPI
    '=====================================
    MPTester.Print "Begin Check MIPI........"
    
    ItemResult = ItemResult + (Check_MIPI * &H8)
    
    If (ItemResult And &H8) = &H8 Then
        MPTester.Print "Check MIPI: PASS"
    Else
        GoTo TestEnd
    End If
    
    '======================================
    '   CSet Value Test
    '======================================
    MPTester.Print "Begin CSet Test........"
    Call MsecDelay(0.1)
    ItemResult = ItemResult + (CSET_Value_Test * &H10)
    
    If (ItemResult And &H10) = &H10 Then
        MPTester.Print "CSet Value: PASS"
    Else
        GoTo TestEnd
    End If
    
TestEnd:
    
    cardresult = DO_WritePort(card, Channel_P1A, &HFB) 'Close ENA Power 1111_1011
    
    If FW_Fail_Flag Then
        FWFail_Counter = FWFail_Counter + 1
    Else
        FWFail_Counter = 0
    End If

    If FWFail_Counter > 10 Then
        MPTester.FWFail_Label.Visible = True
    Else
        MPTester.FWFail_Label.Visible = False
    End If
    
    Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
    
    WaitDevOFF ("058f")
    WaitDevOFF ("058f")
    Call MsecDelay(0.2)
    
    If HV Then
    
        If (ItemResult And &H1) <> &H1 Then
            TestResult = "Bin2"
            MPTester.TestResultLab = "Bin2: HV Set XTAL Fail"
            
        ElseIf (ItemResult And &H40) = &H40 Then
            TestResult = "Bin4"
            MPTester.TestResultLab = "Bin4: HV VDD1812 Condition Pass"
        
        ElseIf (ItemResult And &H80) = &H80 Then
            TestResult = "Bin5"
            MPTester.TestResultLab = "Bin5: HV VDDSASD Condition Pass"
                
        ElseIf (ItemResult And &H20) <> &H20 Then
            TestResult = "Bin2"
            MPTester.TestResultLab = "Bin2: HV LDO Fail"
                
        ElseIf ((ItemResult And &H4) <> &H4) Then
            TestResult = "Bin2"
            MPTester.TestResultLab = "Bin2: HV Parallel Fail"
            
        ElseIf (ItemResult And &H2) <> &H2 Then
            TestResult = "Bin2"
            MPTester.TestResultLab = "Bin2: HV Set LC Fail"
            
        ElseIf (ItemResult And &H8) <> &H8 Then
            TestResult = "Bin3"
            MPTester.TestResultLab = "Bin3: HV MIPI Fail"
           
        ElseIf (ItemResult And &H10) <> &H10 Then
            TestResult = "Bin2"
            MPTester.TestResultLab = "Bin2: HV CSet Fail"
            
        ElseIf (ItemResult = &H3F) Then
            TestResult = "PASS"
            MPTester.TestResultLab = "Bin1: HV PASS"
            HLV_Result = HLV_Result + 1
            
        Else
            TestResult = "Bin2"
            MPTester.TestResultLab = "Bin2: Undefine Fail"
        End If
        
        Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
        HV = False
        GoTo HLV_Routine
        
    Else
    
        If (ItemResult And &H1) <> &H1 Then
            TestResult = "Bin2"
            MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "Bin2: LV Set XTAL Fail"
            
        ElseIf (ItemResult And &H40) = &H40 Then
            TestResult = "Bin4"
            MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "Bin4: LV VDD1812 Condition Pass"
        
        ElseIf (ItemResult And &H80) = &H80 Then
            TestResult = "Bin5"
            MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "Bin5: LV VDDSASD Condition Pass"
                
        ElseIf (ItemResult And &H20) <> &H20 Then
            TestResult = "Bin2"
            MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "Bin2: LV LDO Fail"
                
        ElseIf ((ItemResult And &H4) <> &H4) Then
            TestResult = "Bin2"
            MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "Bin2: LV Parallel Fail"
            
        ElseIf (ItemResult And &H2) <> &H2 Then
            TestResult = "Bin2"
            MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "Bin2: LV Set LC Fail"
            
        ElseIf (ItemResult And &H8) <> &H8 Then
            TestResult = "Bin3"
            MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "Bin3: LV MIPI Fail"
           
        ElseIf (ItemResult And &H10) <> &H10 Then
            TestResult = "Bin2"
            MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "Bin2: LV CSet Fail"
            
        ElseIf (ItemResult = &H3F) Then
            TestResult = "PASS"
            MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "Bin1: LV PASS"
            HLV_Result = HLV_Result + 2
            
        Else
            TestResult = "Bin2"
            MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "Bin2: Undefine Fail"
        End If
    End If
    
    If HLV_Result = 3 Then
        TestResult = "PASS"
    ElseIf HLV_Result = 1 Then
        TestResult = "Bin4"                             '"HV Pass, LV Fail"
    ElseIf HLV_Result = 2 Then
        TestResult = "Bin3"                             '"HV Fail, LV Pass"
    Else
        TestResult = "Bin2"                             '"HV and LV Fail"
    End If
    
    HLV_Result = 0
    
    If TestResult = "Bin2" And FailCloseAP Then
        Call CloseVedioCap
    End If
    
End Sub

' modify to check LDO fail

Public Sub AU3826A81AFS10TestSub()

Dim i As Integer
Dim OldTimer As Long
Dim PassTime As Long
Dim rt2 As Long
Dim LDOValue As Long
Dim mMsg As MSG
Dim TmpStr As String
Dim ItemResult As Byte
Dim GPIO_Value As Long
Dim LDOPASSCount As Integer
Dim GPIOReadVal As Long
Dim GPIO15Result As Byte

Dim VDD1812Secondary As Integer
Dim VDDSASDSecondary As Integer
Dim NA As Integer
ReDim LDOVal(1 To 50) As Long

Dim VDD12_high, VDD12_low, VDD18_high, VDD18_low


    If InStr(1, ChipName, "AFS") <> 0 Then
        FileCopy App.Path & "\CamTest\AU3826A81FTTest_40QFN\3826\VideoCap.ini", App.Path & "\CamTest\AU3826A81FTTest_40QFN\VideoCap.ini"
    Else
        FileCopy App.Path & "\CamTest\AU3826A81FTTest_40QFN\3822\VideoCap.ini", App.Path & "\CamTest\AU3826A81FTTest_40QFN\VideoCap.ini"
    End If

    NA = 0
    TestResult = ""
    ItemResult = 0
    AlcorMPMessage = 0
    GPIO15Result = 0
    FW_Fail_Flag = False
    Check3826VCC5V = False

    If PCI7248InitFinish = 0 Then
        Call PCI7248Exist
    End If

    If Not SetP1CInput_Flag Then
       result = DIO_PortConfig(card, Channel_P1C, INPUT_PORT)
       If result <> 0 Then
           MsgBox " config PCI_P1C as input card fail"
           End
       End If
       SetP1CInput_Flag = True
    End If

    cardresult = DO_WritePort(card, Channel_P1A, &HFA)      ' set &HFA to check VCC5V

    If Check3826VCC5V = False Then
        Check3826VCC5V = True
        Call MsecDelay(0.2)
        cardresult = DO_ReadPort(card, Channel_P1B, LDOVal(1))
        If LDOVal(1) = 0 Then
            MsgBox "請確認5V是否有接"
            End
        End If
    End If

    cardresult = DO_WritePort(card, Channel_P1A, &H11)      'to act like open socket behavior
    Call MsecDelay(0.2)

    ' === check device exist and then close power ===
    cardresult = DO_WritePort(card, Channel_P1A, &HEA)      'Open ENA Power 1110_1010 (Bit3 using External clock)
    Call MsecDelay(0.2)
    If Not WaitDevOn("vid_058f") Then
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2:Vid/Pid UnKnow Fail"
        cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'Close ENA Power 1111_1111
        Call CloseVedioCap
        Exit Sub
    End If
    Call MsecDelay(0.7)
    MPTester.TestResultLab = ""

    ' === chipname is new or changed ===
    If OldChipName <> ChipName Then
        ChDir App.Path & "\CamTest\AU3826A81FTTest_40QFN\"

        FileCopy App.Path & "\CamTest\AU3826A81FTTest_40QFN\allow.sys", "C:\WINDOWS\system32\drivers\allow.sys"

        If InStr(1, ChipName, "AFS") <> 0 Then
            Shell App.Path & "\CamTest\AU3826A81FTTest_40QFN\ALInstFtr -d allow 3826"
        Else
            Shell App.Path & "\CamTest\AU3826A81FTTest_40QFN\ALInstFtr -d allow 3822"
        End If

        Call MsecDelay(0.3)

        If InStr(1, ChipName, "AFS") <> 0 Then
            Shell App.Path & "\CamTest\AU3826A81FTTest_40QFN\ALInstFtr -i allow 3826"
        Else
            Shell App.Path & "\CamTest\AU3826A81FTTest_40QFN\ALInstFtr -i allow 3822"
        End If

        Call MsecDelay(0.3)

        FWFail_Counter = 0

        Call CloseVedioCap
        OldChipName = ChipName

        Call Load_VerifyFW_Tool_AU3826A81FTTest_40QFN

        Call MsecDelay(0.5)

        KillProcess ("MPTool_lite_v3.12.620.exe")
        KillProcess ("VerifyFW_v3.2.4.exe")
    End If

    If FindWindow(vbNullString, "VideoCap") = 0 Then
        MPTester.Print "wait for VideoCap Ready"

        OldTimer = Timer

        If LoadVedioCap_AU3826A81FTTest_40QFN Then
            MPTester.Print "Ready Time="; Timer - OldTimer
        Else
            MPTester.TestResultLab = "Bin2:VideoCap Ready Fail "
            TestResult = "Bin2"
            MPTester.Print "VideoCap Ready Fail"
            cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
            Call CloseVedioCap
            Exit Sub
       End If

    End If

    ' ======================== Test result ===========================
    '                        sav/sdv   nc   gpio15   masclk  power
    '                          LDO    EEP     ENC     ENB     ENA
    '   8       7       6       5      4       3       2       1
    '   1       1       1       0      1       0       1       0    check LDO, parallel
    '   1       1       1       1      1       1       0       0    check MIPI
    ' ================================================================

    ' ======================== ItemResult ===========================
    '(2^7) VDDSASD
    '(2^6) VDD1812
    '(2^5) LDO
    '(2^4) CSET
    '
    '(2^3) MIPI
    '(2^2) Parallel
    '(2^1) GPIO15 (LC)
    '(2^0) GPIO15 (XTAL)


    '=====================================
    '   GPIO15 Setting (set to 0) Clock source is XTAL
    '=====================================
    MPTester.Print "Begin GPIO Setting Test........"

    Do
        GPIOReadVal = GPIO_Read(&H21, 0)    'Check GPIO Value is 1
        If (GPIOReadVal = 0) Then
            Exit Do
        End If
        NA = NA + 1
    Loop While (NA < 5)

    NA = 0

    If (GPIOReadVal <> 0) Then
        MPTester.Print "Set XTAL Fail"
        GPIO15Result = 0
        GoTo TestEnd
    Else
        ItemResult = ItemResult + 1
    End If

    '=====================================
    '   Check LDO
    '=====================================
    Call MsecDelay(0.2)
    MPTester.Print "Begin LDO Test........"

'    Call GPIO_Setting(&H560, &H0)
'    Call MsecDelay(0.1)

    ' first do 50 on/off test VDD12 LDO
    For i = 1 To 50

        Call GPIO_Setting(&H560, &H0)
        Call MsecDelay(0.1)

        cardresult = DO_ReadPort(card, Channel_P1B, LDOVal(i))
        Call MsecDelay(0.01)

        VDD12_high = LDOVal(i) And &H80
        VDD12_low = LDOVal(i) And &H40
        VDD18_high = LDOVal(i) And &H20
        VDD18_low = LDOVal(i) And &H10

        If (VDD12_high) Or (VDD12_low) Or (VDD18_high) Or (VDD18_low) Then
            VDD1812Secondary = VDD1812Secondary + 1
            ItemResult = ItemResult + &H80
            MPTester.Print "LDO Fail"
            GoTo SKIP_REAL_LDO_TEST
        End If

        cardresult = DO_WritePort(card, Channel_P1A, &HFF)     'Close ENA Power 1111_1111
        Call MsecDelay(0.05)
        cardresult = DO_WritePort(card, Channel_P1A, &HEA)      'Open ENA Power 1110_1010 (Bit3 using External clock)
        Call MsecDelay(0.05)

    Next

    ' do real LDO test
    Call GPIO_Setting(&H560, &H0)
    Call MsecDelay(1.5)

    For i = 1 To 50
        cardresult = DO_ReadPort(card, Channel_P1B, LDOVal(i))
        Call MsecDelay(0.01)

        If (LDOVal(i) And &H1 = &H1) Or (LDOVal(i) And &H2 = &H2) Or (LDOVal(i) And &H4 = &H4) Or (LDOVal(i) And &H8 = &H8) Then
            VDDSASDSecondary = VDDSASDSecondary + 1
        ElseIf (LDOVal(i) And &H10 = &H10) Or (LDOVal(i) And &H20 = &H20) Or (LDOVal(i) And &H40 = &H40) Or (LDOVal(i) And &H80 = &H80) Then
            VDD1812Secondary = VDD1812Secondary + 1
        Else
            LDOPASSCount = LDOPASSCount + 1
        End If
    Next

    If LDOPASSCount = 50 Then
        MPTester.Print "LDO PASS"
        ItemResult = ItemResult + &H20
    ElseIf (VDDSASDSecondary > 1) And (VDD1812Secondary = 0) Then
        MPTester.Print "VDD1812Secondary pass"
        ItemResult = ItemResult + &H40
    ElseIf (VDDSASDSecondary = 0) And (VDD1812Secondary > 1) Then
        MPTester.Print "VDDSASDSecondary pass"
        ItemResult = ItemResult + &H80
    Else
        MPTester.Print "LDO Fail"
    End If

SKIP_REAL_LDO_TEST:

    '=====================================
    '   Check parallel
    '=====================================
    MPTester.Print "Begin Check Parallel........"

    ItemResult = ItemResult + (Check_Parallel * 4)

    If (ItemResult And &H4) = &H4 Then
        MPTester.Print "Check parallel: PASS"
    Else
        GoTo TestEnd
    End If

    ' ===== close power and open =====
    cardresult = DO_WritePort(card, Channel_P1A, &HF1)      'Close ENA Power 1111_0001
    WaitDevOFF ("058f")
    WaitDevOFF ("058f")
    Call MsecDelay(0.5)
    ' ================================

    cardresult = DO_WritePort(card, Channel_P1A, &HF0)      'Fine tune for stable the test
    Call MsecDelay(0.2)

    '=====================================
    '   Connect MIPI Power(Bit5 =>H) & MSACLK(Bit2 =>L)
    '=====================================
    cardresult = DO_WritePort(card, Channel_P1A, &HFC)      'Open ENA Power 1111_1100 (Bit3 using Internal LC)
    WaitDevOn ("058f")
    Call MsecDelay(0.2)

    '=====================================
    '   GPIO15 Setting (set to 1) Clock source is LC
    '=====================================
    MPTester.Print "Begin GPIO Setting Test........"

    Do
        GPIOReadVal = GPIO_Read(&H21, &HFF)    'Check GPIO Value is 1
        If (GPIOReadVal = &H80) Then
            Exit Do
        End If
        NA = NA + 1
        Call MsecDelay(0.2)
    Loop While (NA < 5)

    NA = 0

    If (GPIOReadVal <> &H80) Then
        MPTester.Print "Set LC Fail"
        GPIO15Result = 0
        GoTo TestEnd
    Else
        ItemResult = ItemResult + 2
    End If

    '=====================================
    '   Check MIPI
    '=====================================
    MPTester.Print "Begin Check MIPI........"

    ItemResult = ItemResult + (Check_MIPI * &H8)

    If (ItemResult And &H8) = &H8 Then
        MPTester.Print "Check MIPI: PASS"
    Else
        GoTo TestEnd
    End If

    '======================================
    '   CSet Value Test
    '======================================
    MPTester.Print "Begin CSet Test........"
    Call MsecDelay(0.1)
    ItemResult = ItemResult + (CSET_Value_Test * &H10)

    If (ItemResult And &H10) = &H10 Then
        MPTester.Print "CSet Value: PASS"
    Else
        GoTo TestEnd
    End If

TestEnd:

    cardresult = DO_WritePort(card, Channel_P1A, &HFB) 'Close ENA Power 1111_1011

    If FW_Fail_Flag Then
        FWFail_Counter = FWFail_Counter + 1
    Else
        FWFail_Counter = 0
    End If

    If FWFail_Counter > 10 Then
        MPTester.FWFail_Label.Visible = True
    Else
        MPTester.FWFail_Label.Visible = False
    End If

    WaitDevOFF ("058f")
    WaitDevOFF ("058f")
    Call MsecDelay(0.2)

    If (ItemResult And &H1) <> &H1 Then
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2: Set XTAL Fail"

    ElseIf (ItemResult And &H40) = &H40 Then
        TestResult = "Bin4"
        MPTester.TestResultLab = "Bin4: VDD1812 Condition Pass"

    ElseIf (ItemResult And &H80) = &H80 Then
        TestResult = "Bin5"
        MPTester.TestResultLab = "Bin5: VDDSASD Condition Pass"

    ElseIf (ItemResult And &H20) <> &H20 Then
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2: LDO Fail"

    ElseIf ((ItemResult And &H4) <> &H4) Then
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2: Parallel Fail"

    ElseIf (ItemResult And &H2) <> &H2 Then
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2: Set LC Fail"

    ElseIf (ItemResult And &H8) <> &H8 Then
        TestResult = "Bin3"
        MPTester.TestResultLab = "Bin3: MIPI Fail"

    ElseIf (ItemResult And &H10) <> &H10 Then
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2: CSet Fail"

    ElseIf (ItemResult = &H3F) Then
        TestResult = "PASS"
        MPTester.TestResultLab = "Bin1: PASS"

    Else
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2: Undefine Fail"
    End If

    If TestResult = "Bin2" And FailCloseAP Then
        Call CloseVedioCap
    End If

End Sub

' sorting only for Vdd12 fail (XTAL)

Public Sub AU3826A81AFS20TestSub()

Dim i As Integer
Dim OldTimer As Long
Dim PassTime As Long
Dim rt2 As Long
Dim LDOValue As Long
Dim mMsg As MSG
Dim TmpStr As String
Dim ItemResult As Byte
Dim GPIO_Value As Long
Dim LDOPASSCount As Integer
Dim GPIOReadVal As Long
Dim GPIO15Result As Byte

Dim VDD1812Secondary As Integer
Dim VDDSASDSecondary As Integer
Dim NA As Integer
ReDim LDOVal(1 To 50) As Long

Dim VDD12_high, VDD12_low, VDD18_high, VDD18_low

    If (InStr(1, ChipName, "AFS") <> 0 Or InStr(1, ChipName, "CFS") <> 0 Or InStr(1, ChipName, "DFE") <> 0 Or InStr(1, ChipName, "DFS") <> 0) Then
        FileCopy App.Path & "\CamTest\AU3826A81FTTest_40QFN\3826\VideoCap.ini", App.Path & "\CamTest\AU3826A81FTTest_40QFN\VideoCap.ini"
    Else
        FileCopy App.Path & "\CamTest\AU3826A81FTTest_40QFN\3822\VideoCap.ini", App.Path & "\CamTest\AU3826A81FTTest_40QFN\VideoCap.ini"
    End If

    NA = 0
    TestResult = ""
    ItemResult = 0
    AlcorMPMessage = 0
    GPIO15Result = 0
    FW_Fail_Flag = False
    Check3826VCC5V = False

    If PCI7248InitFinish = 0 Then
        Call PCI7248Exist
    End If

    If Not SetP1CInput_Flag Then
       result = DIO_PortConfig(card, Channel_P1C, INPUT_PORT)
       If result <> 0 Then
           MsgBox " config PCI_P1C as input card fail"
           End
       End If
       SetP1CInput_Flag = True
    End If

    cardresult = DO_WritePort(card, Channel_P1A, &HCA)      ' set &HFA to check VCC5V

    If Check3826VCC5V = False Then
        Check3826VCC5V = True
        Call MsecDelay(0.2)
        cardresult = DO_ReadPort(card, Channel_P1B, LDOVal(1))
        If LDOVal(1) = 0 Then
            MsgBox "請確認5V是否有接"
            End
        End If
    End If

    cardresult = DO_WritePort(card, Channel_P1A, &H11)      'to act like open socket behavior
    Call MsecDelay(0.2)

    ' === check device exist and then close power ===
    cardresult = DO_WritePort(card, Channel_P1A, &HCA)      'Open ENA Power 1110_1010 (Bit3 using External clock)
    Call MsecDelay(0.2)
    If Not WaitDevOn("vid_058f") Then
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2:Vid/Pid UnKnow Fail"
        cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'Close ENA Power 1111_1111
        Call CloseVedioCap
        Exit Sub
    End If
    Call MsecDelay(0.7)
    MPTester.TestResultLab = ""

    ' === chipname is new or changed ===
    If OldChipName <> ChipName Then
        ChDir App.Path & "\CamTest\AU3826A81FTTest_40QFN\"

        FileCopy App.Path & "\CamTest\AU3826A81FTTest_40QFN\allow.sys", "C:\WINDOWS\system32\drivers\allow.sys"

        If (InStr(1, ChipName, "AFS") <> 0 Or InStr(1, ChipName, "CFS") <> 0 Or InStr(1, ChipName, "DFE") <> 0 Or InStr(1, ChipName, "DFS") <> 0) Then
            Shell App.Path & "\CamTest\AU3826A81FTTest_40QFN\ALInstFtr -d allow 3826"
        Else
            Shell App.Path & "\CamTest\AU3826A81FTTest_40QFN\ALInstFtr -d allow 3822"
        End If

        Call MsecDelay(0.3)

        If (InStr(1, ChipName, "AFS") <> 0 Or InStr(1, ChipName, "CFS") <> 0 Or InStr(1, ChipName, "DFE") <> 0 Or InStr(1, ChipName, "DFS") <> 0) Then
            Shell App.Path & "\CamTest\AU3826A81FTTest_40QFN\ALInstFtr -i allow 3826"
        Else
            Shell App.Path & "\CamTest\AU3826A81FTTest_40QFN\ALInstFtr -i allow 3822"
        End If

        Call MsecDelay(0.3)

        FWFail_Counter = 0

        Call CloseVedioCap
        OldChipName = ChipName

        Call Load_VerifyFW_Tool_AU3826A81FTTest_40QFN

        Call MsecDelay(0.5)

        KillProcess ("MPTool_lite_v3.12.620.exe")
        KillProcess ("VerifyFW_v3.2.4.exe")
    End If

    '=====================================
    '   Check LDO
    '=====================================
    Call MsecDelay(0.2)
    MPTester.Print "Begin LDO Test........"


    ' first do 50 on/off test VDD12 LDO
    For i = 1 To 50

        Call GPIO_Setting(&H560, &H0)
        Call MsecDelay(0.1)

        cardresult = DO_ReadPort(card, Channel_P1B, LDOVal(i))
        Call MsecDelay(0.01)

        VDD12_high = LDOVal(i) And &H80
        VDD12_low = LDOVal(i) And &H40
        VDD18_high = LDOVal(i) And &H20
        VDD18_low = LDOVal(i) And &H10

        If (VDD12_high) Or (VDD12_low) Or (VDD18_high) Or (VDD18_low) Then
            VDD1812Secondary = VDD1812Secondary + 1
            GoTo TestEnd
        Else
            LDOPASSCount = LDOPASSCount + 1
        End If
       
        If InStr(1, ChipName, "22") <> 0 Then
            cardresult = DO_WritePort(card, Channel_P1A, &HF1)  'Close ENA Power 1111_0001
        Else
            cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'Close ENA Power 1111_1111
        End If
        Call MsecDelay(0.05)
        
        cardresult = DO_WritePort(card, Channel_P1A, &HCA)      'Open ENA Power 1110_1010 (Bit3 using External clock)
        Call MsecDelay(0.05)

    Next

    If LDOPASSCount = 50 Then
        MPTester.Print "LDO PASS"
        ItemResult = ItemResult + &H20
    Else
        MPTester.Print "LDO Fail"
        ItemResult = ItemResult + &H40
    End If

TestEnd:

    If InStr(1, ChipName, "22") <> 0 Then
        cardresult = DO_WritePort(card, Channel_P1A, &HF1) 'Close ENA Power 1111_0001
    Else
        cardresult = DO_WritePort(card, Channel_P1A, &HFB) 'Close ENA Power 1111_1011
    End If

    If FW_Fail_Flag Then
        FWFail_Counter = FWFail_Counter + 1
    Else
        FWFail_Counter = 0
    End If

    If FWFail_Counter > 10 Then
        MPTester.FWFail_Label.Visible = True
    Else
        MPTester.FWFail_Label.Visible = False
    End If

    WaitDevOFF ("058f")
    WaitDevOFF ("058f")
    Call MsecDelay(0.2)


    If (ItemResult = &H20) Then
        TestResult = "PASS"
        MPTester.TestResultLab = "Bin1: PASS"

    Else
        TestResult = "Bin3"
        MPTester.TestResultLab = "Bin3: LDO fail"
    End If

End Sub

' sorting only for Vdd12 fail (LC)

Public Sub AU3826A81AFS24TestSub()

Dim i As Integer
Dim OldTimer As Long
Dim PassTime As Long
Dim rt2 As Long
Dim LDOValue As Long
Dim mMsg As MSG
Dim TmpStr As String
Dim ItemResult As Byte
Dim GPIO_Value As Long
Dim LDOPASSCount As Integer
Dim GPIOReadVal As Long
Dim GPIO15Result As Byte

Dim VDD1812Secondary As Integer
Dim VDDSASDSecondary As Integer
Dim NA As Integer
ReDim LDOVal(1 To 50) As Long

Dim VDD12_high, VDD12_low, VDD18_high, VDD18_low

    If (InStr(1, ChipName, "AFS") <> 0 Or InStr(1, ChipName, "CFS") <> 0 Or InStr(1, ChipName, "DFE") <> 0 Or InStr(1, ChipName, "DFS") <> 0) Then
        FileCopy App.Path & "\CamTest\AU3826A81FTTest_40QFN\3826\VideoCap.ini", App.Path & "\CamTest\AU3826A81FTTest_40QFN\VideoCap.ini"
    Else
        FileCopy App.Path & "\CamTest\AU3826A81FTTest_40QFN\3822\VideoCap.ini", App.Path & "\CamTest\AU3826A81FTTest_40QFN\VideoCap.ini"
    End If

    NA = 0
    TestResult = ""
    ItemResult = 0
    AlcorMPMessage = 0
    GPIO15Result = 0
    FW_Fail_Flag = False
    Check3826VCC5V = False

    If PCI7248InitFinish = 0 Then
        Call PCI7248Exist
    End If

    If Not SetP1CInput_Flag Then
       result = DIO_PortConfig(card, Channel_P1C, INPUT_PORT)
       If result <> 0 Then
           MsgBox " config PCI_P1C as input card fail"
           End
       End If
       SetP1CInput_Flag = True
    End If

    cardresult = DO_WritePort(card, Channel_P1A, &HCA)      ' set &HFA to check VCC5V

    If Check3826VCC5V = False Then
        Check3826VCC5V = True
        Call MsecDelay(0.2)
        cardresult = DO_ReadPort(card, Channel_P1B, LDOVal(1))
        If LDOVal(1) = 0 Then
            MsgBox "請確認5V是否有接"
            End
        End If
    End If

    cardresult = DO_WritePort(card, Channel_P1A, &H11)      'to act like open socket behavior
    Call MsecDelay(0.2)

    ' === check device exist and then close power ===
    cardresult = DO_WritePort(card, Channel_P1A, &HCA)      'Open ENA Power 1110_1010 (Bit3 using External clock)
    
    Call MsecDelay(0.2)
    If Not WaitDevOn("vid_058f") Then
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2:Vid/Pid UnKnow Fail"
        cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'Close ENA Power 1111_1111
        Call CloseVedioCap
        Exit Sub
    End If
    Call MsecDelay(0.7)
    MPTester.TestResultLab = ""

    ' === chipname is new or changed ===
    If OldChipName <> ChipName Then
        ChDir App.Path & "\CamTest\AU3826A81FTTest_40QFN\"

        FileCopy App.Path & "\CamTest\AU3826A81FTTest_40QFN\allow.sys", "C:\WINDOWS\system32\drivers\allow.sys"

        If (InStr(1, ChipName, "AFS") <> 0 Or InStr(1, ChipName, "CFS") <> 0 Or InStr(1, ChipName, "DFE") <> 0 Or InStr(1, ChipName, "DFS") <> 0) Then
            Shell App.Path & "\CamTest\AU3826A81FTTest_40QFN\ALInstFtr -d allow 3826"
        Else
            Shell App.Path & "\CamTest\AU3826A81FTTest_40QFN\ALInstFtr -d allow 3822"
        End If

        Call MsecDelay(0.3)

        If (InStr(1, ChipName, "AFS") <> 0 Or InStr(1, ChipName, "CFS") <> 0 Or InStr(1, ChipName, "DFE") <> 0 Or InStr(1, ChipName, "DFS") <> 0) Then
            Shell App.Path & "\CamTest\AU3826A81FTTest_40QFN\ALInstFtr -i allow 3826"
        Else
            Shell App.Path & "\CamTest\AU3826A81FTTest_40QFN\ALInstFtr -i allow 3822"
        End If

        Call MsecDelay(0.3)

        FWFail_Counter = 0

        Call CloseVedioCap
        OldChipName = ChipName

        Call Load_VerifyFW_Tool_AU3826A81FTTest_40QFN

        Call MsecDelay(0.5)

        KillProcess ("MPTool_lite_v3.12.620.exe")
        KillProcess ("VerifyFW_v3.2.4.exe")
    End If
    
    
    ' 20140505 add for test LDO in LC mode
    '=====================================
    '   GPIO15 Setting (set to 1) Clock source is LC
    '=====================================
    MPTester.Print "Begin GPIO Setting Test........"
    
    ' ===== close power and open =====
    cardresult = DO_WritePort(card, Channel_P1A, &HD1)      'Close ENA Power 1111_0001
    WaitDevOFF ("058f")
    WaitDevOFF ("058f")
    Call MsecDelay(0.05)
    ' ================================

    cardresult = DO_WritePort(card, Channel_P1A, &HCF)      ' set LC mode first
    Call MsecDelay(0.05)
    cardresult = DO_WritePort(card, Channel_P1A, &HCE)      ' then ena on
    Call MsecDelay(0.05)


    '=====================================
    '   Check LDO
    '=====================================
    Call MsecDelay(0.2)
    MPTester.Print "Begin LDO Test........"


    ' first do 50 on/off test VDD12 LDO
    For i = 1 To 50

        Call GPIO_Setting(&H560, &H0)
        Call MsecDelay(0.1)

        cardresult = DO_ReadPort(card, Channel_P1B, LDOVal(i))
        Call MsecDelay(0.01)

        VDD12_high = LDOVal(i) And &H80
        VDD12_low = LDOVal(i) And &H40
        VDD18_high = LDOVal(i) And &H20
        VDD18_low = LDOVal(i) And &H10

        If (VDD12_high) Or (VDD12_low) Or (VDD18_high) Or (VDD18_low) Then
            VDD1812Secondary = VDD1812Secondary + 1
            GoTo TestEnd
        Else
            LDOPASSCount = LDOPASSCount + 1
        End If
       
        If InStr(1, ChipName, "22") <> 0 Then
            cardresult = DO_WritePort(card, Channel_P1A, &HF1)  'Close ENA Power 1111_0001
        Else
            cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'Close ENA Power 1111_1111
        End If
        Call MsecDelay(0.05)
        
        cardresult = DO_WritePort(card, Channel_P1A, &HCE)      'Open ENA Power 1111_1100 (Bit3 using Internal LC)
        
        Call MsecDelay(0.05)

    Next

    If LDOPASSCount = 50 Then
        MPTester.Print "LDO PASS"
        ItemResult = ItemResult + &H20
    Else
        MPTester.Print "LDO Fail"
        ItemResult = ItemResult + &H40
    End If

TestEnd:

    If InStr(1, ChipName, "22") <> 0 Then
        cardresult = DO_WritePort(card, Channel_P1A, &HF1) 'Close ENA Power 1111_0001
    Else
        cardresult = DO_WritePort(card, Channel_P1A, &HFB) 'Close ENA Power 1111_1011
    End If

    If FW_Fail_Flag Then
        FWFail_Counter = FWFail_Counter + 1
    Else
        FWFail_Counter = 0
    End If

    If FWFail_Counter > 10 Then
        MPTester.FWFail_Label.Visible = True
    Else
        MPTester.FWFail_Label.Visible = False
    End If

    WaitDevOFF ("058f")
    WaitDevOFF ("058f")
    Call MsecDelay(0.2)


    If (ItemResult = &H20) Then
        TestResult = "PASS"
        MPTester.TestResultLab = "Bin1: PASS"

    Else
        TestResult = "Bin3"
        MPTester.TestResultLab = "Bin3: LDO fail"
    End If

End Sub

Public Sub AU3825D61BFF2GTestSub()
'add unload driver function

If PCI7248InitFinish = 0 Then
      Call PCI7248Exist
End If

If Not SetP1CInput_Flag Then
   result = DIO_PortConfig(card, Channel_P1C, INPUT_PORT)
   If result <> 0 Then
       MsgBox " config PCI_P1C as input card fail"
       End
   End If
   SetP1CInput_Flag = True
End If

Dim i As Integer
Dim OldTimer As Long
Dim PassTime As Long
Dim rt2 As Long
Dim LDOValue As Long
Dim mMsg As MSG
Dim LDORetry As Byte
Dim TmpStr As String
Dim TempResult As Byte
Dim GPIO_Value As Long
Dim TempCount As Byte
Dim SRAMPASSCount As Byte
Dim V18FailCount As Integer
Dim SecondaryCount As Integer
Dim LDOPASSCount As Integer
Dim GPIOReadVal As Long
Dim GPIO15Result As Byte
ReDim LDOVal(1 To 50) As Long

TestResult = ""
TempResult = 0
LDORetry = 0
AlcorMPMessage = 0
GPIO15Result = 0
FW_Fail_Flag = False

cardresult = DO_WritePort(card, Channel_P1A, &HFA) 'Open ENA Power 1111_1010 (Bit3 using External clock)

Call MsecDelay(0.2)
If Not WaitDevOn("vid_058f") Then
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2:Vid/Pid UnKnow Fail"
    cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
    Call CloseVedioCap
    Exit Sub
End If
Call MsecDelay(0.8)
MPTester.TestResultLab = ""
'===============================================================
' Fail location initial
'===============================================================

If OldChipName <> ChipName Then

    ChDir App.Path & "\CamTest\AU3825A61FTTest_40QFN\"

    If Dir("C:\WINDOWS\system32\drivers\allow.sys") = "allow.sys" Then
        Kill ("C:\WINDOWS\system32\drivers\allow.sys")
    End If

    FWFail_Counter = 0

    Call CloseVedioCap
    OldChipName = ChipName

    Call Load_VerifyFW_Tool_AU3825_40QFN

    KillProcess ("VerifyFW_v3.2.4.exe")
End If

If FindWindow(vbNullString, "VideoCap") = 0 Then

    MPTester.Print "wait for VideoCap Ready"

    OldTimer = Timer

    If LoadVedioCap_AU3825_40QFN Then
        MPTester.Print "Ready Time="; Timer - OldTimer
    Else
        MPTester.TestResultLab = "Bin2:VideoCap Ready Fail "
        TestResult = "Bin2"
        MPTester.Print "VideoCap Ready Fail"
        cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
        Call CloseVedioCap
        Exit Sub
   End If
End If


'TempResult
'Bit 1,2  : GPIO_Setting        11: PASS
'Bit 3    : Image pattern        1: PASS
'Bit 4,5  : LDO                 11: PASS, 01:Condition PASS
'Bit 6    : SRAM                 1: PASS
'Bit 7,8    : CSET Value        01: PASS

'=====================================
'   Set ST1 Level3
'=====================================
'Call GPIO_Setting(&H564, &H64)
'Call MsecDelay(0.02)
'Call GPIO_Setting(&H52A, &HB)
'Call MsecDelay(0.02)
'Call GPIO_Setting(&H52B, &H78)
'Call MsecDelay(0.02)


'=====================================
'   GPIO Setting
'=====================================
MPTester.Print "Begin GPIO Setting Test........"

TempCount = 0

Do
    Call GPIO_Setting(&H20, &H54)
    Call MsecDelay(0.04)
    cardresult = DO_ReadPort(card, Channel_P1C, GPIO_Value)
    Call MsecDelay(0.02)

    If GPIO_Value = &HE9 Then
        TempResult = TempResult + 1
        Exit Do
    End If

    TempCount = TempCount + 1
    Call MsecDelay(0.02)

Loop Until (TempCount > 10)


TempCount = 0
Call MsecDelay(0.02)
Do
    Call GPIO_Setting(&H20, &H29)
    Call MsecDelay(0.04)
    cardresult = DO_ReadPort(card, Channel_P1C, GPIO_Value)
    Call MsecDelay(0.02)

    If GPIO_Value = &HD6 Then
        TempResult = TempResult + 2
        Exit Do
    End If

    TempCount = TempCount + 1
    Call MsecDelay(0.02)
Loop Until (TempCount > 10)


If (TempResult And &H3) = 3 Then

    GPIOReadVal = GPIO_Read(&H21, 0)
    If (GPIOReadVal And &H40) <> 0 Then     'Xtal: 0x21 bit7 must be "L"
        MPTester.Print "GPIO15 Fail"
        GPIO15Result = 0
        GoTo TestEnd
    End If

    MPTester.Print "GPIO Setting: PASS"
Else
    GoTo TestEnd
End If


'=====================================
'   LDO
'=====================================
Call MsecDelay(0.2)
MPTester.Print "Begin LDO Test........"

For i = 1 To 50
    cardresult = DO_ReadPort(card, Channel_P1B, LDOVal(i))
    Call MsecDelay(0.01)
Next

For i = 1 To 50
    If (LDOVal(i) And &H3) <> 0 Then                            'VDD18 Fail
        V18FailCount = V18FailCount + 1
    ElseIf (LDOVal(i) = &H40) Or (LDOVal(i) = &H20) Then    'XSA < 2.65 or XSA > 2.95
        SecondaryCount = SecondaryCount + 1
    ElseIf LDOVal(i) = 0 Then
        LDOPASSCount = LDOPASSCount + 1
    End If
Next

'MPTester.Print "V18: " & V18FailCount
'MPTester.Print "Condition: " & SecondaryCount
'MPTester.Print "LDO PSS: " & LDOPASSCount

If V18FailCount >= 1 Then
    MPTester.Print "VDD18 Fail"
ElseIf (SecondaryCount >= 1) And (V18FailCount = 0) Then
    MPTester.Print "LDO Condition PASS"
    TempResult = TempResult + &H8
ElseIf LDOPASSCount = 50 Then
    MPTester.Print "LDO PASS"
    TempResult = TempResult + &H18
Else
    MPTester.Print "LDO Fail"
End If

'If LDOValue = 0 Then
'    MPTester.Print "LDO PASS"
'    TempResult = TempResult + &H18
'ElseIf (LDOValue = &H40) Or (LDOValue = &H20) Then      'XSA < 2.65 or XSA > 2.95
'    MPTester.Print "LDO Condition PASS"
'    TempResult = TempResult + &H8
'ElseIf (LDOValue And &H3) <> 0 Then                     'VDD18 Fail
'    MPTester.Print "Bin2: VDD18 Fail"
'    TestResult = "Bin2"
'    MPTester.TestResultLab = "Bin2: VDD18 Fail"
'    cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'Close ENA Power 1111_1111
'    TempCounter = TempCounter + 1
'    TempVDD18Counter = TempVDD18Counter + 1
'    Exit Sub
'Else
'    MPTester.Print "LDO Fail"
'End If

'Call MsecDelay(0.2)

'=====================================
'   Image
'=====================================
MPTester.Print "Begin Image Test........"
'OldTimer = Timer

TempResult = TempResult + (Image_Test * &H4)

If (TempResult And &H4) = &H4 Then
    MPTester.Print "Image Test: PASS"
Else
    GoTo TestEnd
End If

'======================================
'   Set LV & SRAM Test
'======================================

Call MsecDelay(0.3)
cardresult = DO_WritePort(card, Channel_P1A, &HF8) 'Select External、Enable External 1.62V、Open ENA Power XTAL 1111_1000
Call MsecDelay(0.8)

For SRAMPASSCount = 1 To 2
    If SRAM_Test = 0 Then
        MPTester.Print "SRAM Test " & " Cycle " & SRAMPASSCount & ": Fail"
        Exit For
    Else
        MPTester.Print "Cycle " & SRAMPASSCount & ": PASS"
    End If

    If SRAMPASSCount = 2 Then
        TempResult = TempResult + (SRAM_Test * &H20)

        If (TempResult And &H20) = &H20 Then
            MPTester.Print "SRAM Test: PASS"
        End If
    End If
Next

If (TempResult And &H20) <> &H20 Then
    GoTo TestEnd
End If

cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Power OFF
WaitDevOFF ("vid_058f")
Call MsecDelay(0.2)
cardresult = DO_WritePort(card, Channel_P1A, &HFE) 'Power ON (Ena、Internal RC)
Call MsecDelay(0.2)


If Not WaitDevOn("vid_058f") Then
    TestResult = "Bin4"
    MPTester.TestResultLab = "Bin4:Internal UnKnow Fail"
    cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
    Exit Sub
End If
Call MsecDelay(0.4)

'======================================
'   CSet Value Test
'======================================
MPTester.Print "Begin CSet Test........"
Call MsecDelay(0.1)
TempResult = TempResult + (CSET_Value_Test * &H40)

If winHwnd = FindWindow(vbNullString, "MPTool_lite_v3.12.620") Then

    If winHwnd <> 0 Then
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call CloseVedioCap
        KillProcess ("MPTool_lite_v3.12.620.exe")
        'Call LoadVedioCap_AU3825_40QFN
        'Call MsecDelay(1#)
        'TempResult = TempResult + (CSET_Value_Test * &H40)
    End If
End If

If (TempResult And &HC0) = &H40 Then
    MPTester.Print "CSet Value: PASS"
End If
Call MsecDelay(0.02)

GPIOReadVal = 0
GPIOReadVal = GPIO_Read(&H21, 0)
If (CByte(GPIOReadVal) And &H80) <> &H80 Then    'RC 0x21 bit7 must be "H"
    MPTester.Print "GPIO15 Fail"
    GPIO15Result = 0
Else
    GPIO15Result = 1
End If

TestEnd:

cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111

If FW_Fail_Flag Then
    FWFail_Counter = FWFail_Counter + 1
Else
    FWFail_Counter = 0
End If

If FWFail_Counter > 10 Then
    MPTester.FWFail_Label.Visible = True
Else
    MPTester.FWFail_Label.Visible = False
End If

WaitDevOFF ("058f")
WaitDevOFF ("058f")
Call MsecDelay(0.2)

If (TempResult And &H3) <> &H3 Then
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2: GPIO Setting Fail"
    'TempGPIOCounter = TempGPIOCounter + 1

ElseIf (GPIO15Result <> 1) Then
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2: GPIO15 Fail"

ElseIf ((TempResult And &H4) <> &H4) Then
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2: Image Fail"
    'TempImageCounter = TempImageCounter + 1

ElseIf (TempResult And &H20) <> &H20 Then
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2: SRAM Fail"
    'TempSRAMCounter = TempSRAMCounter + 1

ElseIf (TempResult And &HC0) <> &H40 Then
    TestResult = "Bin4"
    MPTester.TestResultLab = "Bin4: CSet Fail"
    'TempCSETCounter = TempCSETCounter + 1

ElseIf (V18FailCount <> 0) Then
    TestResult = "Bin4"
    MPTester.TestResultLab = "Bin4: VDD18 Fail"

ElseIf (TempResult And &H18) = &H0 Then
    TestResult = "Bin3"
    MPTester.TestResultLab = "Bin3: LDO Fail"
    'TempLDOCounter = TempLDOCounter + 1

ElseIf (TempResult And &H18) = &H8 Then
    TestResult = "Bin5"
    MPTester.TestResultLab = "Bin5: Secondary PASS"
    'TempConditionCounter = TempConditionCounter + 1
ElseIf (TempResult = &H7F) Then
    TestResult = "PASS"
    MPTester.TestResultLab = "Bin1: PASS"
    'TempPASSCounter = TempPASSCounter + 1
Else
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2: Undefine Fail"
End If

If TestResult = "Bin2" And FailCloseAP Then
    Call CloseVedioCap
End If

'TempCounter = TempCounter + 1
'If TempCounter = 30 Then
'    'debug.print "PASS: " & TempPASSCounter & " ;SRAM: " & TempSRAMCounter & " ;CSET: "; TempCSETCounter _
'                ; " ;GPIO: " & TempGPIOCounter & " ;Condition: " & TempConditionCounter & " ;Image: " & TempImageCounter _
'                ; " ;LDO: " & TempLDOCounter & " ;VDD18: " & TempVDD18Counter
'
'    TempCounter = 0
'    TempSRAMCounter = 0
'    TempCSETCounter = 0
'    TempGPIOCounter = 0
'    TempConditionCounter = 0
'    TempPASSCounter = 0
'    TempImageCounter = 0
'    TempLDOCounter = 0
'    TempVDD18Counter = 0
'
'End If

End Sub

Public Sub AU3825D61BFE10TestSub()
'add unload driver function

If PCI7248InitFinish = 0 Then
      Call PCI7248Exist
End If

If Not SetP1CInput_Flag Then
   result = DIO_PortConfig(card, Channel_P1C, INPUT_PORT)
   If result <> 0 Then
       MsgBox " config PCI_P1C as input card fail"
       End
   End If
   SetP1CInput_Flag = True
End If

Dim i As Integer
Dim OldTimer As Long
Dim PassTime As Long
Dim rt2 As Long
Dim LDOValue As Long
Dim mMsg As MSG
Dim LDORetry As Byte
Dim TmpStr As String
Dim TempResult As Byte
Dim GPIO_Value As Long
Dim TempCount As Byte
Dim SRAMPASSCount As Byte
Dim V18FailCount As Integer
Dim SecondaryCount As Integer
Dim LDOPASSCount As Integer
Dim SCFailCount As Integer

Dim GPIOReadVal As Long
Dim GPIO15Result As Byte
ReDim LDOVal(1 To 50) As Long

    TestResult = ""
    TempResult = 0
    LDORetry = 0
    AlcorMPMessage = 0
    GPIO15Result = 0
    FW_Fail_Flag = False

    cardresult = DO_WritePort(card, Channel_P1A, &HFA) 'Open ENA Power 1111_1010 (Bit3 using External clock)


    '=====================================
    '   LDO
    '=====================================
    Call MsecDelay(0.2)
    MPTester.Print "Begin LDO Test........"

    For i = 1 To 50
        cardresult = DO_ReadPort(card, Channel_P1B, LDOVal(i))
        Call MsecDelay(0.01)
    Next

    For i = 1 To 50
        If (LDOVal(i) And &H3) <> 0 Then                            'VDD18 Fail
            V18FailCount = V18FailCount + 1
        ElseIf (LDOVal(i) And &HC) <> 0 Then
            SCFailCount = SCFailCount + 1
        ElseIf (LDOVal(i) And &HF0) <> 0 Then
            SecondaryCount = SecondaryCount + 1
        ElseIf LDOVal(i) = 0 Then
            LDOPASSCount = LDOPASSCount + 1
        End If
    Next
    
    If V18FailCount >= 1 Then
        MPTester.Print "VDD18 Fail"
    ElseIf (SCFailCount >= 1) And (V18FailCount = 0) Then
        MPTester.Print "SC Fail"
        'TempResult = TempResult + &H8
    ElseIf (SecondaryCount >= 1) And (SCFailCount = 0) Then
        MPTester.Print "SA Fail"
        'TempResult = TempResult + &H10
    ElseIf LDOPASSCount = 50 Then
        MPTester.Print "LDO PASS"
        TempResult = TempResult + &H18
    Else
        MPTester.Print "LDO Fail"
    End If

TestEnd:

    cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111

    If (V18FailCount <> 0) Then
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2: VDD18 Fail"
    ElseIf (SCFailCount <> 0) Then
        TestResult = "Bin3"
        MPTester.TestResultLab = "Bin3 : XSC Fail"
    ElseIf (SecondaryCount <> 0) Then
        TestResult = "Bin5"
        MPTester.TestResultLab = "Bin5 : XSA Fail"
    ElseIf (TempResult = &H18) Then
        TestResult = "PASS"
        MPTester.TestResultLab = "Bin1: PASS"
    Else
        TestResult = "Bin2"
        MPTester.TestResultLab = "Bin2: Undefine Fail"
    End If

End Sub

Public Sub AU3825A61BFImgTestSub()
'add unload driver function

If PCI7248InitFinish = 0 Then
      Call PCI7248Exist
End If

If Not SetP1CInput_Flag Then
   result = DIO_PortConfig(card, Channel_P1C, INPUT_PORT)
   If result <> 0 Then
       MsgBox " config PCI_P1C as input card fail"
       End
   End If
   SetP1CInput_Flag = True
End If
 
Dim i As Integer
Dim OldTimer As Long
Dim PassTime As Long
Dim rt2 As Long
Dim LDOValue As Long
Dim mMsg As MSG
Dim LDORetry As Byte
Dim TmpStr As String
Dim TempResult As Byte
Dim GPIO_Value As Long
Dim TempCount As Byte
Dim SRAMPASSCount As Byte
Dim V18FailCount As Integer
Dim SecondaryCount As Integer
Dim LDOPASSCount As Integer
ReDim LDOVal(1 To 50) As Long
Dim RC_ImageResult As Byte
 
TestResult = ""
TempResult = 0
LDORetry = 0
AlcorMPMessage = 0
FW_Fail_Flag = False

Call PowerSet2(1, "3.1", "0.5", 1, "3.1", "0.5", 1)
cardresult = DO_WritePort(card, Channel_P1A, &HFA) 'Open ENA Power 1111_1010 (Bit3 using External clock)

Call MsecDelay(0.2)
If Not WaitDevOn("vid_058f") Then
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2:Vid/Pid UnKnow Fail"
    cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
    Call CloseVedioCap
    Exit Sub
End If
Call MsecDelay(0.8)
MPTester.TestResultLab = ""
'===============================================================
' Fail location initial
'===============================================================

If OldChipName <> ChipName Then
    
    ChDir App.Path & "\CamTest\AU3825A61FTTest_40QFN\"
    
    If Dir("C:\WINDOWS\system32\drivers\allow.sys") = "allow.sys" Then
        Kill ("C:\WINDOWS\system32\drivers\allow.sys")
    End If
    
    FWFail_Counter = 0
    
    Call CloseVedioCap
    OldChipName = ChipName

    Call Load_VerifyFW_Tool_AU3825_40QFN
    
    KillProcess ("VerifyFW_v3.2.4.exe")
End If

If FindWindow(vbNullString, "VideoCap") = 0 Then
    
    MPTester.Print "wait for VideoCap Ready"
    
    OldTimer = Timer
    
    If LoadVedioCap_AU3825_40QFN Then
        MPTester.Print "Ready Time="; Timer - OldTimer
    Else
        MPTester.TestResultLab = "Bin2:VideoCap Ready Fail "
        TestResult = "Bin2"
        MPTester.Print "VideoCap Ready Fail"
        cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
        Call CloseVedioCap
        Exit Sub
   End If
End If
         

'TempResult
'Bit 1,2  : GPIO_Setting        11: PASS
'Bit 3    : Image pattern        1: PASS
'Bit 4,5  : LDO                 11: PASS, 01:Condition PASS
'Bit 6    : SRAM                 1: PASS
'Bit 7,8    : CSET Value        01: PASS
                 
'=====================================
'   Set ST1 Level3
'=====================================
'Call GPIO_Setting(&H564, &H64)
'Call MsecDelay(0.02)
'Call GPIO_Setting(&H52A, &HB)
'Call MsecDelay(0.02)
'Call GPIO_Setting(&H52B, &H78)
'Call MsecDelay(0.02)


'=====================================
'   GPIO Setting
'=====================================
MPTester.Print "Begin GPIO Setting Test........"

TempCount = 0

Do
    Call GPIO_Setting(&H20, &H54)
    Call MsecDelay(0.04)
    cardresult = DO_ReadPort(card, Channel_P1C, GPIO_Value)
    Call MsecDelay(0.02)

    If GPIO_Value = &HE9 Then
        TempResult = TempResult + 1
        Exit Do
    End If
    
    TempCount = TempCount + 1
    Call MsecDelay(0.02)
    
Loop Until (TempCount > 10)


TempCount = 0
Call MsecDelay(0.02)
Do
    Call GPIO_Setting(&H20, &H29)
    Call MsecDelay(0.04)
    cardresult = DO_ReadPort(card, Channel_P1C, GPIO_Value)
    Call MsecDelay(0.02)
    
    If GPIO_Value = &HD6 Then
        TempResult = TempResult + 2
        Exit Do
    End If
    
    TempCount = TempCount + 1
    Call MsecDelay(0.02)
Loop Until (TempCount > 10)


If (TempResult And &H3) = 3 Then
    MPTester.Print "GPIO Setting: PASS"
Else
    GoTo TestEnd
End If


'=====================================
'   LDO
'=====================================
'Call MsecDelay(0.2)
'MPTester.Print "Begin LDO Test........"
'
'For i = 1 To 50
'    cardresult = DO_ReadPort(card, Channel_P1B, LDOVal(i))
'    Call MsecDelay(0.01)
'Next
'
'For i = 1 To 50
'    If (LDOVal(i) And &H3) <> 0 Then                            'VDD18 Fail
'        V18FailCount = V18FailCount + 1
'    ElseIf (LDOVal(i) = &H40) Or (LDOVal(i) = &H20) Then    'XSA < 2.65 or XSA > 2.95
'        SecondaryCount = SecondaryCount + 1
'    ElseIf LDOVal(i) = 0 Then
'        LDOPASSCount = LDOPASSCount + 1
'    End If
'Next

TempResult = TempResult + &H18   'bypass LDO test item

'MPTester.Print "V18: " & V18FailCount
'MPTester.Print "Condition: " & SecondaryCount
'MPTester.Print "LDO PSS: " & LDOPASSCount

'If V18FailCount >= 1 Then
'    MPTester.Print "VDD18 Fail"
'ElseIf (SecondaryCount >= 1) And (V18FailCount = 0) Then
'    MPTester.Print "LDO Condition PASS"
'    TempResult = TempResult + &H8
'ElseIf LDOPASSCount = 50 Then
'    MPTester.Print "LDO PASS"
'    TempResult = TempResult + &H18
'Else
'    MPTester.Print "LDO Fail"
'End If

'If LDOValue = 0 Then
'    MPTester.Print "LDO PASS"
'    TempResult = TempResult + &H18
'ElseIf (LDOValue = &H40) Or (LDOValue = &H20) Then      'XSA < 2.65 or XSA > 2.95
'    MPTester.Print "LDO Condition PASS"
'    TempResult = TempResult + &H8
'ElseIf (LDOValue And &H3) <> 0 Then                     'VDD18 Fail
'    MPTester.Print "Bin2: VDD18 Fail"
'    TestResult = "Bin2"
'    MPTester.TestResultLab = "Bin2: VDD18 Fail"
'    cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'Close ENA Power 1111_1111
'    TempCounter = TempCounter + 1
'    TempVDD18Counter = TempVDD18Counter + 1
'    Exit Sub
'Else
'    MPTester.Print "LDO Fail"
'End If

'Call MsecDelay(0.2)

'=====================================
'   Image
'=====================================
MPTester.Print "Begin Image Test........"
'OldTimer = Timer

TempResult = TempResult + (Image_Test * &H4)

If (TempResult And &H4) = &H4 Then
    MPTester.Print "Image Test: PASS"
Else
    GoTo TestEnd
End If

'======================================
'   Set LV & SRAM Test
'======================================

Call MsecDelay(0.3)
cardresult = DO_WritePort(card, Channel_P1A, &HF8) 'Select External、Enable External 1.62V、Open ENA Power XTAL 1111_1000
Call MsecDelay(0.8)

For SRAMPASSCount = 1 To 2
    If SRAM_Test = 0 Then
        MPTester.Print "SRAM Test " & " Cycle " & SRAMPASSCount & ": Fail"
        Exit For
    Else
        MPTester.Print "Cycle " & SRAMPASSCount & ": PASS"
    End If
    
    If SRAMPASSCount = 2 Then
        TempResult = TempResult + (SRAM_Test * &H20)

        If (TempResult And &H20) = &H20 Then
            MPTester.Print "SRAM Test: PASS"
        End If
    End If
Next

If (TempResult And &H20) <> &H20 Then
    GoTo TestEnd
End If

cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Power OFF
WaitDevOFF ("vid_058f")
Call MsecDelay(0.2)
cardresult = DO_WritePort(card, Channel_P1A, &HFE) 'Power ON (Ena、Internal RC)
Call MsecDelay(0.2)


If Not WaitDevOn("vid_058f") Then
    TestResult = "Bin4"
    MPTester.TestResultLab = "Bin4:Internal UnKnow Fail"
    cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
    Exit Sub
End If
Call MsecDelay(0.8)

'======================================
'   RC Image Test
'======================================
MPTester.Print "Begin RC Image Test........"
RC_ImageResult = Image_Test
If RC_ImageResult <> 1 Then
    MPTester.Print "RC Image Test: Fail"
Else
    MPTester.Print "RC Image Test: PASS"
End If

'======================================
'   CSet Value Test
'======================================
MPTester.Print "Begin CSet Test........"
Call MsecDelay(0.1)
TempResult = TempResult + (CSET_Value_Test * &H40)

If (TempResult And &HC0) = &H40 Then
    MPTester.Print "CSet Value: PASS"
End If


If winHwnd = FindWindow(vbNullString, "MPTool_lite_v3.12.620") Then
        
    If winHwnd <> 0 Then
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call CloseVedioCap
        KillProcess ("MPTool_lite_v3.12.620.exe")
        'Call LoadVedioCap_AU3825_40QFN
        'Call MsecDelay(1#)
        'TempResult = TempResult + (CSET_Value_Test * &H40)
    End If
End If


TestEnd:

cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111

If FW_Fail_Flag Then
    FWFail_Counter = FWFail_Counter + 1
Else
    FWFail_Counter = 0
End If

If FWFail_Counter > 10 Then
    MPTester.FWFail_Label.Visible = True
Else
    MPTester.FWFail_Label.Visible = False
End If

WaitDevOFF ("058f")
WaitDevOFF ("058f")
Call MsecDelay(0.2)

If (TempResult And &H3) <> &H3 Then
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2: GPIO Setting Fail"
    'TempGPIOCounter = TempGPIOCounter + 1
ElseIf ((TempResult And &H4) <> &H4) Then
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2: Image Fail"
    'TempImageCounter = TempImageCounter + 1
    
ElseIf (TempResult And &H20) <> &H20 Then
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2: SRAM Fail"
    'TempSRAMCounter = TempSRAMCounter + 1
    
ElseIf (TempResult And &HC0) <> &H40 Then
    TestResult = "Bin4"
    MPTester.TestResultLab = "Bin4: CSet Fail"
    'TempCSETCounter = TempCSETCounter + 1

ElseIf RC_ImageResult <> 1 Then
    TestResult = "Bin5"
    MPTester.TestResultLab = "Bin5: RC Image Fail"

ElseIf (V18FailCount <> 0) Then
    TestResult = "Bin4"
    MPTester.TestResultLab = "Bin4: VDD18 Fail"

'ElseIf (TempResult And &H18) = &H0 Then
'    TestResult = "Bin3"
'    MPTester.TestResultLab = "Bin3: LDO Fail"
    'TempLDOCounter = TempLDOCounter + 1
    
'ElseIf (TempResult And &H18) = &H8 Then
'    TestResult = "Bin5"
'    MPTester.TestResultLab = "Bin5: Secondary PASS"
    'TempConditionCounter = TempConditionCounter + 1
ElseIf (TempResult = &H7F) Then
    TestResult = "PASS"
    MPTester.TestResultLab = "Bin1: PASS"
    'TempPASSCounter = TempPASSCounter + 1
Else
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2: Undefine Fail"
End If
                            
If TestResult = "Bin2" And FailCloseAP Then
    Call CloseVedioCap
End If
                            
'TempCounter = TempCounter + 1
'If TempCounter = 30 Then
'    'debug.print "PASS: " & TempPASSCounter & " ;SRAM: " & TempSRAMCounter & " ;CSET: "; TempCSETCounter _
'                ; " ;GPIO: " & TempGPIOCounter & " ;Condition: " & TempConditionCounter & " ;Image: " & TempImageCounter _
'                ; " ;LDO: " & TempLDOCounter & " ;VDD18: " & TempVDD18Counter
'
'    TempCounter = 0
'    TempSRAMCounter = 0
'    TempCSETCounter = 0
'    TempGPIOCounter = 0
'    TempConditionCounter = 0
'    TempPASSCounter = 0
'    TempImageCounter = 0
'    TempLDOCounter = 0
'    TempVDD18Counter = 0
'
'End If
                            
End Sub

Public Sub AU3825A61BFS7ETestSub()
'add unload driver function

'2012/8/30: base on ST1, skip LDO test & enhance SRAM test 7 cycle

If PCI7248InitFinish = 0 Then
      Call PCI7248Exist
End If

If Not SetP1CInput_Flag Then
   result = DIO_PortConfig(card, Channel_P1C, INPUT_PORT)
   If result <> 0 Then
       MsgBox " config PCI_P1C as input card fail"
       End
   End If
   SetP1CInput_Flag = True
End If
 
Dim i As Integer
Dim OldTimer As Long
Dim PassTime As Long
Dim rt2 As Long
Dim LDOValue As Long
Dim mMsg As MSG
Dim LDORetry As Byte
Dim TmpStr As String
Dim TempResult As Byte
Dim GPIO_Value As Long
Dim TempCount As Byte
Dim SRAMPASSCount As Byte
Dim V18FailCount As Integer
Dim SecondaryCount As Integer
Dim LDOPASSCount As Integer
ReDim LDOVal(1 To 50) As Long
 
TestResult = ""
TempResult = 0
LDORetry = 0
AlcorMPMessage = 0
FW_Fail_Flag = False
 
cardresult = DO_WritePort(card, Channel_P1A, &HFA) 'Open ENA Power 1111_1010 (Bit3 using External clock)

Call MsecDelay(0.2)
If Not WaitDevOn("vid_058f") Then
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2:Vid/Pid UnKnow Fail"
    cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
    Call CloseVedioCap
    Exit Sub
End If
Call MsecDelay(0.8)
MPTester.TestResultLab = ""
'===============================================================
' Fail location initial
'===============================================================

If OldChipName <> ChipName Then
    
    ChDir App.Path & "\CamTest\AU3825A61FTTest_40QFN\"
    
    If Dir("C:\WINDOWS\system32\drivers\allow.sys") = "allow.sys" Then
        Kill ("C:\WINDOWS\system32\drivers\allow.sys")
    End If
    
    FWFail_Counter = 0
    
    Call CloseVedioCap
    OldChipName = ChipName

    Call Load_VerifyFW_Tool_AU3825_40QFN
    
    KillProcess ("VerifyFW_v3.2.4.exe")
End If

If FindWindow(vbNullString, "VideoCap") = 0 Then
    
    MPTester.Print "wait for VideoCap Ready"
    
    OldTimer = Timer
    
    If LoadVedioCap_AU3825_40QFN Then
        MPTester.Print "Ready Time="; Timer - OldTimer
    Else
        MPTester.TestResultLab = "Bin2:VideoCap Ready Fail "
        TestResult = "Bin2"
        MPTester.Print "VideoCap Ready Fail"
        cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
        Call CloseVedioCap
        Exit Sub
   End If
End If
         

'TempResult
'Bit 1,2  : GPIO_Setting        11: PASS
'Bit 3    : Image pattern        1: PASS
'Bit 4,5  : LDO                 11: PASS, 01:Condition PASS
'Bit 6    : SRAM                 1: PASS
'Bit 7,8    : CSET Value        01: PASS

'=====================================
'   GPIO Setting
'=====================================
MPTester.Print "Begin GPIO Setting Test........"

TempCount = 0

Do
    Call GPIO_Setting(&H20, &H54)
    Call MsecDelay(0.04)
    cardresult = DO_ReadPort(card, Channel_P1C, GPIO_Value)
    Call MsecDelay(0.02)

    If GPIO_Value = &HE9 Then
        TempResult = TempResult + 1
        Exit Do
    End If
    
    TempCount = TempCount + 1
    Call MsecDelay(0.02)
    
Loop Until (TempCount > 10)


TempCount = 0
Call MsecDelay(0.02)
Do
    Call GPIO_Setting(&H20, &H29)
    Call MsecDelay(0.04)
    cardresult = DO_ReadPort(card, Channel_P1C, GPIO_Value)
    Call MsecDelay(0.02)
    
    If GPIO_Value = &HD6 Then
        TempResult = TempResult + 2
        Exit Do
    End If
    
    TempCount = TempCount + 1
    Call MsecDelay(0.02)
Loop Until (TempCount > 10)


If (TempResult And &H3) = 3 Then
    MPTester.Print "GPIO Setting: PASS"
Else
    GoTo TestEnd
End If

Call MsecDelay(0.3)
'=====================================
'   Image
'=====================================
MPTester.Print "Begin Image Test........"
'OldTimer = Timer

TempResult = TempResult + (Image_Test * &H4)

If (TempResult And &H4) = &H4 Then
    MPTester.Print "Image Test: PASS"
Else
    GoTo TestEnd
End If

'======================================
'   Set LV & SRAM Test
'======================================

Call MsecDelay(0.3)
cardresult = DO_WritePort(card, Channel_P1A, &HF8) 'Select External、Enable External 1.62V、Open ENA Power XTAL 1111_1000
Call MsecDelay(0.8)

For SRAMPASSCount = 1 To 7
    If SRAM_Test = 0 Then
        MPTester.Print "SRAM Test " & " Cycle " & SRAMPASSCount & ": Fail"
        Exit For
    Else
        MPTester.Print "Cycle " & SRAMPASSCount & ": PASS"
    End If
    
    If SRAMPASSCount = 7 Then
        TempResult = TempResult + (SRAM_Test * &H20)

        If (TempResult And &H20) = &H20 Then
            MPTester.Print "SRAM Test: PASS"
        End If
    End If
Next

If (TempResult And &H20) <> &H20 Then
    GoTo TestEnd
End If

cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Power OFF
WaitDevOFF ("vid_058f")
Call MsecDelay(0.2)
cardresult = DO_WritePort(card, Channel_P1A, &HFE) 'Power ON (Ena、Internal RC)
Call MsecDelay(0.2)


If Not WaitDevOn("vid_058f") Then
    TestResult = "Bin4"
    MPTester.TestResultLab = "Bin4:Internal UnKnow Fail"
    cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
    Exit Sub
End If
Call MsecDelay(0.4)

'======================================
'   CSet Value Test
'======================================
MPTester.Print "Begin CSet Test........"
Call MsecDelay(0.1)
TempResult = TempResult + (CSET_Value_Test * &H40)

If winHwnd = FindWindow(vbNullString, "MPTool_lite_v3.12.620") Then
        
    If winHwnd <> 0 Then
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call CloseVedioCap
        KillProcess ("MPTool_lite_v3.12.620.exe")
        'Call LoadVedioCap_AU3825_40QFN
        'Call MsecDelay(1#)
        'TempResult = TempResult + (CSET_Value_Test * &H40)
    End If
End If

If (TempResult And &HC0) = &H40 Then
    MPTester.Print "CSet Value: PASS"
End If
Call MsecDelay(0.02)


TestEnd:

cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111

If FW_Fail_Flag Then
    FWFail_Counter = FWFail_Counter + 1
Else
    FWFail_Counter = 0
End If

If FWFail_Counter > 10 Then
    MPTester.FWFail_Label.Visible = True
Else
    MPTester.FWFail_Label.Visible = False
End If

WaitDevOFF ("058f")
WaitDevOFF ("058f")
Call MsecDelay(0.2)

If (TempResult And &H3) <> &H3 Then
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2: GPIO Setting Fail"
    'TempGPIOCounter = TempGPIOCounter + 1
ElseIf ((TempResult And &H4) <> &H4) Then
    TestResult = "Bin3"
    MPTester.TestResultLab = "Bin3: Image Fail"
    'TempImageCounter = TempImageCounter + 1
    
ElseIf (TempResult And &H20) <> &H20 Then
    TestResult = "Bin5"
    MPTester.TestResultLab = "Bin5: SRAM Fail"
    'TempSRAMCounter = TempSRAMCounter + 1
    
ElseIf (TempResult And &HC0) <> &H40 Then
    TestResult = "Bin4"
    MPTester.TestResultLab = "Bin4: CSet Fail"
    'TempCSETCounter = TempCSETCounter + 1

ElseIf (TempResult = &H67) Then
    TestResult = "PASS"
    MPTester.TestResultLab = "Bin1: PASS"
    'TempPASSCounter = TempPASSCounter + 1
Else
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2: Undefine Fail"
End If
                            
If TestResult = "Bin2" And FailCloseAP Then
    Call CloseVedioCap
End If

                            
End Sub

Public Sub AU3825A61BFS6ETestSub()
'add unload driver function
'2012/8/27: This code copy from AU3825A61BFF2E
'           purpose to sorting bin2 all fail binning
'
'           Bin2: Unknow
'           Bin3: GPIO、Image
'           Bin4: VDD18
'           Bin5: SRAM


If PCI7248InitFinish = 0 Then
      Call PCI7248Exist
End If

If Not SetP1CInput_Flag Then
   result = DIO_PortConfig(card, Channel_P1C, INPUT_PORT)
   If result <> 0 Then
       MsgBox " config PCI_P1C as input card fail"
       End
   End If
   SetP1CInput_Flag = True
End If
 
Dim i As Integer
Dim OldTimer As Long
Dim PassTime As Long
Dim rt2 As Long
Dim LDOValue As Long
Dim mMsg As MSG
Dim LDORetry As Byte
Dim TmpStr As String
Dim TempResult As Byte
Dim GPIO_Value As Long
Dim TempCount As Byte
Dim SRAMPASSCount As Byte
Dim V18FailCount As Integer
Dim SecondaryCount As Integer
Dim LDOPASSCount As Integer
ReDim LDOVal(1 To 50) As Long
 
TestResult = ""
TempResult = 0
LDORetry = 0
AlcorMPMessage = 0
FW_Fail_Flag = False
 
cardresult = DO_WritePort(card, Channel_P1A, &HFA) 'Open ENA Power 1111_1010 (Bit3 using External clock)

Call MsecDelay(0.2)
If Not WaitDevOn("vid_058f") Then
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2:Vid/Pid UnKnow Fail"
    cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
    Call CloseVedioCap
    Exit Sub
End If
Call MsecDelay(0.8)
MPTester.TestResultLab = ""
'===============================================================
' Fail location initial
'===============================================================

If OldChipName <> ChipName Then
    
    ChDir App.Path & "\CamTest\AU3825A61FTTest_40QFN\"
    
    If Dir("C:\WINDOWS\system32\drivers\allow.sys") = "allow.sys" Then
        Kill ("C:\WINDOWS\system32\drivers\allow.sys")
    End If
    
    FWFail_Counter = 0
    
    Call CloseVedioCap
    OldChipName = ChipName

    Call Load_VerifyFW_Tool_AU3825_40QFN
    
    KillProcess ("VerifyFW_v3.2.4.exe")
End If

If FindWindow(vbNullString, "VideoCap") = 0 Then
    
    MPTester.Print "wait for VideoCap Ready"
    
    OldTimer = Timer
    
    If LoadVedioCap_AU3825_40QFN Then
        MPTester.Print "Ready Time="; Timer - OldTimer
    Else
        MPTester.TestResultLab = "Bin2:VideoCap Ready Fail "
        TestResult = "Bin2"
        MPTester.Print "VideoCap Ready Fail"
        cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
        Call CloseVedioCap
        Exit Sub
   End If
End If
         

'TempResult
'Bit 1,2  : GPIO_Setting        11: PASS
'Bit 3    : Image pattern        1: PASS
'Bit 4,5  : LDO                 11: PASS, 01:Condition PASS
'Bit 6    : SRAM                 1: PASS
'Bit 7,8    : CSET Value        01: PASS
                 

'=====================================
'   GPIO Setting
'=====================================
MPTester.Print "Begin GPIO Setting Test........"

TempCount = 0

Do
    Call GPIO_Setting(&H20, &H54)
    Call MsecDelay(0.04)
    cardresult = DO_ReadPort(card, Channel_P1C, GPIO_Value)
    Call MsecDelay(0.02)

    If GPIO_Value = &HE9 Then
        TempResult = TempResult + 1
        Exit Do
    End If
    
    TempCount = TempCount + 1
    Call MsecDelay(0.02)
    
Loop Until (TempCount > 10)


TempCount = 0
Call MsecDelay(0.02)
Do
    Call GPIO_Setting(&H20, &H29)
    Call MsecDelay(0.04)
    cardresult = DO_ReadPort(card, Channel_P1C, GPIO_Value)
    Call MsecDelay(0.02)
    
    If GPIO_Value = &HD6 Then
        TempResult = TempResult + 2
        Exit Do
    End If
    
    TempCount = TempCount + 1
    Call MsecDelay(0.02)
Loop Until (TempCount > 10)


If (TempResult And &H3) = 3 Then
    MPTester.Print "GPIO Setting: PASS"
Else
    GoTo TestEnd
End If


'=====================================
'   Image
'=====================================
MPTester.Print "Begin Image Test........"
'OldTimer = Timer

TempResult = TempResult + (Image_Test * &H4)

If (TempResult And &H4) = &H4 Then
    MPTester.Print "Image Test: PASS"
Else
    GoTo TestEnd
End If


'=====================================
'   LDO
'=====================================
Call MsecDelay(0.2)
MPTester.Print "Begin LDO Test........"

For i = 1 To 50
    cardresult = DO_ReadPort(card, Channel_P1B, LDOVal(i))
    Call MsecDelay(0.01)
Next

For i = 1 To 50
    If (LDOVal(i) And &H3) <> 0 Then                            'VDD18 Fail
        V18FailCount = V18FailCount + 1
    ElseIf (LDOVal(i) = &H40) Or (LDOVal(i) = &H20) Then    'XSA < 2.65 or XSA > 2.95
        SecondaryCount = SecondaryCount + 1
    ElseIf LDOVal(i) = 0 Then
        LDOPASSCount = LDOPASSCount + 1
    End If
Next

'MPTester.Print "V18: " & V18FailCount
'MPTester.Print "Condition: " & SecondaryCount
'MPTester.Print "LDO PSS: " & LDOPASSCount

If V18FailCount >= 1 Then
    MPTester.Print "Bin4: VDD18 Fail"
    TestResult = "Bin4"
    MPTester.TestResultLab = "Bin4: VDD18 Fail"
    cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'Close ENA Power 1111_1111
    'TempCounter = TempCounter + 1
    'TempVDD18Counter = TempVDD18Counter + 1
    Call CloseVedioCap
    Exit Sub
'ElseIf SecondaryCount >= 1 Then
'    MPTester.Print "LDO Condition PASS"
'    TempResult = TempResult + &H8
'ElseIf LDOPASSCount = 50 Then
'    MPTester.Print "LDO PASS"
'    TempResult = TempResult + &H18
'Else
'    MPTester.Print "LDO Fail"
End If

'If LDOValue = 0 Then
'    MPTester.Print "LDO PASS"
'    TempResult = TempResult + &H18
'ElseIf (LDOValue = &H40) Or (LDOValue = &H20) Then      'XSA < 2.65 or XSA > 2.95
'    MPTester.Print "LDO Condition PASS"
'    TempResult = TempResult + &H8
'ElseIf (LDOValue And &H3) <> 0 Then                     'VDD18 Fail
'    MPTester.Print "Bin2: VDD18 Fail"
'    TestResult = "Bin2"
'    MPTester.TestResultLab = "Bin2: VDD18 Fail"
'    cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'Close ENA Power 1111_1111
'    TempCounter = TempCounter + 1
'    TempVDD18Counter = TempVDD18Counter + 1
'    Exit Sub
'Else
'    MPTester.Print "LDO Fail"
'End If

'Call MsecDelay(0.2)

'======================================
'   Set LV & SRAM Test
'======================================

Call MsecDelay(0.3)
cardresult = DO_WritePort(card, Channel_P1A, &HF8) 'Select External、Enable External 1.62V、Open ENA Power XTAL 1111_1000
Call MsecDelay(0.8)

For SRAMPASSCount = 1 To 2
    If SRAM_Test = 0 Then
        MPTester.Print "SRAM Test " & " Cycle " & SRAMPASSCount & ": Fail"
        Exit For
    Else
        MPTester.Print "Cycle " & SRAMPASSCount & ": PASS"
    End If
    
    If SRAMPASSCount = 2 Then
        TempResult = TempResult + (SRAM_Test * &H20)

        If (TempResult And &H20) = &H20 Then
            MPTester.Print "SRAM Test: PASS"
        End If
    End If
Next

If (TempResult And &H20) <> &H20 Then
    GoTo TestEnd
End If


TestEnd:

cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111

If FW_Fail_Flag Then
    FWFail_Counter = FWFail_Counter + 1
Else
    FWFail_Counter = 0
End If

If FWFail_Counter > 10 Then
    MPTester.FWFail_Label.Visible = True
Else
    MPTester.FWFail_Label.Visible = False
End If

WaitDevOFF ("058f")
WaitDevOFF ("058f")
Call MsecDelay(0.2)

If (TempResult And &H3) <> &H3 Then
    TestResult = "Bin3"
    MPTester.TestResultLab = "Bin3: GPIO Setting Fail"
    'TempGPIOCounter = TempGPIOCounter + 1
ElseIf ((TempResult And &H4) <> &H4) Then
    TestResult = "Bin3"
    MPTester.TestResultLab = "Bin3: Image Fail"
    'TempImageCounter = TempImageCounter + 1
    
ElseIf (TempResult And &H20) <> &H20 Then
    TestResult = "Bin5"
    MPTester.TestResultLab = "Bin5: SRAM Fail"
    'TempSRAMCounter = TempSRAMCounter + 1
    
'ElseIf (TempResult And &HC0) <> &H40 Then
'    TestResult = "Bin4"
'    MPTester.TestResultLab = "Bin4: CSet Fail"
    'TempCSETCounter = TempCSETCounter + 1

'ElseIf (TempResult And &H18) = &H0 Then
'    TestResult = "Bin3"
'    MPTester.TestResultLab = "Bin3: LDO Fail"
    'TempLDOCounter = TempLDOCounter + 1
    
'ElseIf (TempResult And &H18) = &H8 Then
'    TestResult = "Bin5"
'    MPTester.TestResultLab = "Bin5: Secondary PASS"
    'TempConditionCounter = TempConditionCounter + 1
'ElseIf (TempResult = &H7F) Then
'    TestResult = "PASS"
'    MPTester.TestResultLab = "Bin1: PASS"
    'TempPASSCounter = TempPASSCounter + 1
Else
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2: Undefine Fail"
End If
                            
If TestResult = "Bin2" And FailCloseAP Then
    Call CloseVedioCap
End If
                            
'TempCounter = TempCounter + 1
'If TempCounter = 30 Then
'    'debug.print "PASS: " & TempPASSCounter & " ;SRAM: " & TempSRAMCounter & " ;CSET: "; TempCSETCounter _
'                ; " ;GPIO: " & TempGPIOCounter & " ;Condition: " & TempConditionCounter & " ;Image: " & TempImageCounter _
'                ; " ;LDO: " & TempLDOCounter & " ;VDD18: " & TempVDD18Counter
'
'    TempCounter = 0
'    TempSRAMCounter = 0
'    TempCSETCounter = 0
'    TempGPIOCounter = 0
'    TempConditionCounter = 0
'    TempPASSCounter = 0
'    TempImageCounter = 0
'    TempLDOCounter = 0
'    TempVDD18Counter = 0
'
'End If
                            
End Sub

Public Sub AU3825A61SortingTestSub()

'2012/5/31 DBF Use XTAL clock source, just run on Win7 OS
'          DAF Use Resonator

'add unload driver function
If PCI7248InitFinish = 0 Then
    Call PCI7248Exist
End If

If Not SetP1CInput_Flag Then
   result = DIO_PortConfig(card, Channel_P1C, INPUT_PORT)
   If result <> 0 Then
       MsgBox " config PCI_P1C as input card fail"
       End
   End If
   SetP1CInput_Flag = True
End If

Dim i As Integer
Dim OldTimer As Long
Dim PassTime As Long
Dim rt2 As Long
Dim LDOValue As Long
Dim mMsg As MSG
Dim FailCycleLog_W As Long
Dim FailCycleLog_L As Long


If (DriverDieCount >= 3) Then
    Call MsecDelay(2)
    Shell "cmd /c shutdown -r  -t 0", vbHide
End If
    

TestResult = ""
AlcorMPMessage = 0
 
OldTimer = Timer
cardresult = DO_WritePort(card, Channel_P1A, &HFA) 'Open ENA Power 1111_1010 (Bit3 using External clock)

Call MsecDelay(0.2)
If Not WaitDevOn("vid_058f") Then
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2:Vid/Pid UnKnow Fail"
    cardresult = DO_WritePort(card, Channel_P1A, &HFB) 'Close ENA Power 1111_1011
    DriverDieCount = DriverDieCount + 1
    Exit Sub
End If
Call MsecDelay(0.5)
MPTester.TestResultLab = ""
'===============================================================
' Fail location initial
'===============================================================

If OldChipName <> ChipName Then
    
    ChDir App.Path & "\CamTest\AU3825A61FTTest\"
    
    If Dir("C:\WINDOWS\system32\drivers\allow.sys") = "allow.sys" Then
        Kill ("C:\WINDOWS\system32\drivers\allow.sys")
    End If
    
    Call CloseVedioCap
    OldChipName = ChipName
End If

'find window
winHwnd = FindWindow(vbNullString, "VideoCap")
If winHwnd = 0 Then
    
    MPTester.Print "wait for VideoCap Ready"
    
    OldTimer = Timer
     
    Call ShellExecute(MPTester.hwnd, "open", App.Path & "\CamTest\AU3825A61SortTest\VideoCap.exe", "", "", SW_SHOW)
    SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags
    AlcorMPMessage = 0
    
    Do
        ' DoEvents
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
        
        PassTime = Timer - OldTimer
    
    Loop Until (AlcorMPMessage = WM_CAM_MP_READY) Or (PassTime > 5)
    MPTester.Print "AP ReadyTime = " & PassTime
End If
         
If PassTime > 5 Then
    Call CloseVedioCap
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2:AP Ready Fail"
    cardresult = DO_WritePort(card, Channel_P1A, &HFB) 'Close ENA Power 1111_1011
    DriverDieCount = DriverDieCount + 1
    Exit Sub
End If
         
'=====================================
'   Sorting Test
'=====================================
MPTester.Print "Begin Sorting Test........"

winHwnd = FindWindow(vbNullString, "VideoCap")
rt2 = PostMessage(winHwnd, WM_CAM_PHY_START, ByVal 500, ByVal 0)    'lPara = Test Cycle

Do
    If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
        AlcorMPMessage = mMsg.message
        FailCycleLog_W = mMsg.wParam
        FailCycleLog_L = mMsg.lParam
        TranslateMessage mMsg
        DispatchMessage mMsg
    End If
     
    PassTime = Timer - OldTimer
   
Loop Until AlcorMPMessage = WM_CAM_PHY_PASS _
      Or AlcorMPMessage = WM_CAM_PHY_UNKNOW_FAIL _
      Or AlcorMPMessage = WM_CAM_PHY_FAIL _
      Or PassTime > 90
    
TestEnd:

cardresult = DO_WritePort(card, Channel_P1A, &HFB) 'Close ENA Power 1111_1011
WaitDevOFF ("058f")
Call MsecDelay(0.2)

Open "D:\AU3825SortingLog_" & ChipName & "_" & CPUInfo.dwProcessorType & ".txt" For Append As #20

If PassTime > 90 Then
    TestResult = "Bin3"
    Print #20, FailCycleLog_W & " , " & FailCycleLog_L & " (Bin3:Time-Out)"
    MPTester.TestResultLab = "Bin3: Time-Out Fail"
    
ElseIf AlcorMPMessage = WM_CAM_PHY_UNKNOW_FAIL Then
    TestResult = "Bin4"
    Print #20, FailCycleLog_W & " , " & FailCycleLog_L & " (Bin4:UNKNOW)"
    MPTester.TestResultLab = "Bin4: PHY UNKNOW Fail"
    If (FailCycleLog_W = 0) Then
        DriverDieCount = DriverDieCount + 1
    End If
    
ElseIf AlcorMPMessage = WM_CAM_PHY_FAIL Then
    TestResult = "Bin5"
    Print #20, FailCycleLog_W & " , " & FailCycleLog_L & " (Bin5:PHY-FAIL)"
    MPTester.TestResultLab = "Bin5: PHY Fail"
    If (FailCycleLog_W = 0) Then
        DriverDieCount = DriverDieCount + 1
    End If
    
ElseIf AlcorMPMessage = WM_CAM_PHY_PASS Then
    TestResult = "PASS"
    Print #20, FailCycleLog_W & " , " & FailCycleLog_L & " (Bin1:PASS)"
    MPTester.TestResultLab = "Bin1: PASS"
    DriverDieCount = 0
Else
    TestResult = "Bin2"
    Print #20, FailCycleLog_W & " , " & FailCycleLog_L & " (Bin2:Undefine)"
    MPTester.TestResultLab = "Bin2: Undefine Fail"
End If
                            
Close #20
                            
                            
If (TestResult <> "PASS") And (FailCloseAP) Then
    Call CloseVedioCap
End If
                            
                            
End Sub

Public Sub AU3825A61ST4TestSub()

'2012/6/12 AU3825A61-DAF/DBF using phy-board test
'add unload driver function

'config P1C H-byte as input     (H-byte Bit1: Comp0(GPIO0), bit2: Comp1(GPIO1))
'config P1C L-byte as output    (L-byte Bit1: Reset(GPIO4))



If PCI7248InitFinish = 0 Then
    Call PCI7248Exist
End If

If Not SetP1CST4Cond_Flag Then
   result = DIO_PortConfig(card, Channel_P1CH, INPUT_PORT)
    If result <> 0 Then
        MsgBox " config PCI_P1CH as input card fail"
        End
    End If
    
    result = DIO_PortConfig(card, Channel_P1CL, OUTPUT_PORT)
    If result <> 0 Then
        MsgBox " config PCI_P1CL as output card fail"
        End
    End If
    SetP1CST4Cond_Flag = True
End If

Dim i As Integer
Dim OldTimer As Long
Dim PassTime As Long
Dim PhyResult As Long

OldTimer = Timer




'Open S/B Ena
cardresult = DO_WritePort(card, Channel_P1A, &HFA) 'Open ENA Power 1111_1010 (Bit3 using RC clock)
Call MsecDelay(0.2)

cardresult = DO_WritePort(card, Channel_P1A, &H7A) 'Open ENA Power 0111_1010 (Bit3 using RC clock)
Call MsecDelay(0.2)

cardresult = DO_WritePort(card, Channel_P1CL, &HF) 'Set Reset 1111
Call MsecDelay(0.1)

'Set phy-board "Reset" as "L"
cardresult = DO_WritePort(card, Channel_P1CL, &HE) 'Set Reset 1110
Call MsecDelay(0.2)

'Start Test (Release Reset pin)
cardresult = DO_WritePort(card, Channel_P1CL, &HF) 'Set Reset 1111
Call MsecDelay(12#)

cardresult = DO_ReadPort(card, Channel_P1CH, PhyResult)
Call MsecDelay(0.1)



TestEnd:

cardresult = DO_WritePort(card, Channel_P1A, &HFB) 'Close ENA Power 1111_1011
cardresult = DO_WritePort(card, Channel_P1CL, &HE) 'Set Reset 1111


If (PhyResult = PHY_TEST_Fail_1) Or (PhyResult = PHY_TEST_Fail_2) Then
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2: PHY Test Fail"
    
ElseIf PhyResult = PHY_TEST_UNKNOW Then
    TestResult = "Bin4"
    MPTester.TestResultLab = "Bin4: PHY UNKNOW Fail"

ElseIf PhyResult = PHY_TEST_PASS Then
    TestResult = "PASS"
    MPTester.TestResultLab = "Bin1: PASS"
    
Else
    TestResult = "Bin5"
    MPTester.TestResultLab = "Bin5: Test UNKNOW Fail"

End If
                                                        
End Sub

Public Sub AU3825A61AFF2FTestSub()
'add unload driver function
 If PCI7248InitFinish = 0 Then
       Call PCI7248Exist
 End If
 
 If Not SetP1CInput_Flag Then
    result = DIO_PortConfig(card, Channel_P1C, INPUT_PORT)
    If result <> 0 Then
        MsgBox " config PCI_P1C as input card fail"
        End
    End If
    SetP1CInput_Flag = True
 End If
 
 Dim i As Integer
 Dim OldTimer As Long
 Dim PassTime As Long
 Dim rt2 As Long
 Dim LDOValue As Long
 Dim mMsg As MSG
 Dim LDORetry As Byte
 Dim TmpStr As String
 Dim TempResult As Byte
 Dim GPIO_Value As Long
 Dim TempCount As Byte
 Dim SRAMPASSCount As Byte
 Dim V18FailCount As Integer
 Dim SecondaryCount As Integer
 Dim LDOPASSCount As Integer
 ReDim LDOVal(1 To 50) As Long
 
 TestResult = ""
 TempResult = 0
 LDORetry = 0
 AlcorMPMessage = 0
 FW_Fail_Flag = False
 
cardresult = DO_WritePort(card, Channel_P1A, &HFE) 'Open ENA Power 1111_1110

Call MsecDelay(0.2)
If Not WaitDevOn("vid_058f") Then
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2:Vid/Pid UnKnow Fail"
    cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
    Call CloseVedioCap
    Exit Sub
End If
 
Call MsecDelay(0.8)
 
MPTester.TestResultLab = ""
'===============================================================
' Fail location initial
'===============================================================

If OldChipName <> ChipName Then
    
    ChDir App.Path & "\CamTest\AU3825A61FTTest_28QFN\"
    
    If Dir("C:\WINDOWS\system32\drivers\allow.sys") = "allow.sys" Then
        Kill ("C:\WINDOWS\system32\drivers\allow.sys")
    End If
    FWFail_Counter = 0
    Call CloseVedioCap
    OldChipName = ChipName
    
    Call Load_VerifyFW_Tool_AU3825_28QFN
    
    KillProcess ("VerifyFW_v3.2.4.exe")
    
End If

If FindWindow(vbNullString, "VideoCap") = 0 Then
    
    MPTester.Print "wait for VideoCap Ready"
    
    OldTimer = Timer
    
    If LoadVedioCap_AU3825_28QFN Then
        MPTester.Print "Ready Time="; Timer - OldTimer
    Else
        MPTester.TestResultLab = "Bin2:VideoCap Ready Fail "
        TestResult = "Bin2"
        MPTester.Print "VideoCap Ready Fail"
        cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
        Exit Sub
   End If
End If
         

'TempResult
'Bit 1,2  : GPIO_Setting        11: PASS
'Bit 3    : Image pattern        1: PASS
'Bit 4,5  : LDO                 11: PASS, 01:Condition PASS
'Bit 6    : SRAM                 1: PASS
'Bit 7,8    : CSET Value        01: PASS

'=====================================
'   Set ST1 Level3
'=====================================
'Call GPIO_Setting(&H564, &H64)
'Call MsecDelay(0.02)
'Call GPIO_Setting(&H52A, &HB)
'Call MsecDelay(0.02)
'Call GPIO_Setting(&H52B, &H78)
'Call MsecDelay(0.02)


'======================================
'   CSet Value Test
'======================================
MPTester.Print "Begin CSet Test........"
Call MsecDelay(0.04)
TempResult = TempResult + (CSET_Value_Test * &H40)

If (TempResult And &HC0) = &H40 Then
    MPTester.Print "CSet Value: PASS"
End If
Call MsecDelay(0.02)


'=====================================
'   GPIO Setting
'=====================================
Call MsecDelay(0.1)
MPTester.Print "Begin GPIO Setting Test........"

TempCount = 0

Do
    Call GPIO_Setting(&H20, &H54)
    Call MsecDelay(0.2)
    cardresult = DO_ReadPort(card, Channel_P1C, GPIO_Value)
    Call MsecDelay(0.02)

    If GPIO_Value = &HF9 Then
        TempResult = TempResult + 1
        Exit Do
    End If
    
    TempCount = TempCount + 1
    Call MsecDelay(0.02)
    
Loop Until (TempCount > 10)


TempCount = 0

Do
    Call GPIO_Setting(&H20, &H29)
    Call MsecDelay(0.2)
    cardresult = DO_ReadPort(card, Channel_P1C, GPIO_Value)
    Call MsecDelay(0.02)
    
    If GPIO_Value = &HDE Then
        TempResult = TempResult + 2
        Exit Do
    End If
    
    TempCount = TempCount + 1
    Call MsecDelay(0.02)
Loop Until (TempCount > 10)


If (TempResult And &H3) = 3 Then
    MPTester.Print "GPIO Setting: PASS"
Else
    GoTo TestEnd
End If


'=====================================
'   LDO
'=====================================
MPTester.Print "Begin LDO Test........"

For i = 1 To 50
    cardresult = DO_ReadPort(card, Channel_P1B, LDOVal(i))
    Call MsecDelay(0.01)
Next

For i = 1 To 50
    If (LDOVal(i) And &H3) <> 0 Then                            'VDD18 Fail
        V18FailCount = V18FailCount + 1
    ElseIf (LDOVal(i) = &H40) Or (LDOVal(i) = &H20) Then    'XSA < 2.65 or XSA > 2.95
        SecondaryCount = SecondaryCount + 1
    ElseIf LDOVal(i) = 0 Then
        LDOPASSCount = LDOPASSCount + 1
    End If
Next

'MPTester.Print "V18: " & V18FailCount
'MPTester.Print "Condition: " & SecondaryCount
'MPTester.Print "LDO PSS: " & LDOPASSCount

If V18FailCount >= 1 Then
    MPTester.Print "Bin2: VDD18 Fail"
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2: VDD18 Fail"
    cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'Close ENA Power 1111_1111
    'TempCounter = TempCounter + 1
    'TempVDD18Counter = TempVDD18Counter + 1
    Call CloseVedioCap
    Exit Sub
ElseIf SecondaryCount >= 1 Then
    MPTester.Print "LDO Condition PASS"
    TempResult = TempResult + &H8
ElseIf LDOPASSCount = 50 Then
    MPTester.Print "LDO PASS"
    TempResult = TempResult + &H18
Else
    MPTester.Print "LDO Fail"
End If

'If LDOValue = 0 Then
'    MPTester.Print "LDO PASS"
'    TempResult = TempResult + &H18
'ElseIf (LDOValue = &H40) Or (LDOValue = &H20) Then      'XSA < 2.65 or XSA > 2.95
'    MPTester.Print "LDO Condition PASS"
'    TempResult = TempResult + &H8
'ElseIf (LDOValue And &H3) <> 0 Then                     'VDD18 Fail
'    MPTester.Print "Bin2: VDD18 Fail"
'    TestResult = "Bin2"
'    MPTester.TestResultLab = "Bin2: VDD18 Fail"
'    cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'Close ENA Power 1111_1111
'    TempCounter = TempCounter + 1
'    TempVDD18Counter = TempVDD18Counter + 1
'    Exit Sub
'Else
'    MPTester.Print "LDO Fail"
'End If

'=====================================
'   Image
'=====================================
Call MsecDelay(0.2)
MPTester.Print "Begin Image Test........"
'OldTimer = Timer

TempResult = TempResult + (Image_Test * &H4)

If (TempResult And &H4) = &H4 Then
    MPTester.Print "Image Test: PASS"
Else
    GoTo TestEnd
End If

'======================================
'   Set LV & SRAM Test
'======================================

cardresult = DO_WritePort(card, Channel_P1A, &HFC) 'Open ENA Power 1111_1100
Call MsecDelay(0.8)

For SRAMPASSCount = 1 To 2
    If SRAM_Test = 0 Then
        MPTester.Print "SRAM Test " & " Cycle " & SRAMPASSCount & ": Fail"
        Exit For
    Else
        MPTester.Print "Cycle " & SRAMPASSCount & ": PASS"
    End If
    
    If SRAMPASSCount = 2 Then
        TempResult = TempResult + (SRAM_Test * &H20)

        If (TempResult And &H20) = &H20 Then
            MPTester.Print "SRAM Test: PASS"
        End If
    End If
Next


TestEnd:

KillProcess ("VerifyFW_v3.2.4.exe")

cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
WaitDevOFF ("058f")
WaitDevOFF ("058f")
Call MsecDelay(0.2)

If FW_Fail_Flag Then
    FWFail_Counter = FWFail_Counter + 1
Else
    FWFail_Counter = 0
End If

If FWFail_Counter > 10 Then
    MPTester.FWFail_Label.Visible = True
Else
    MPTester.FWFail_Label.Visible = False
End If


If (TempResult And &H3) <> &H3 Then
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2: GPIO Setting Fail"
    'TempGPIOCounter = TempGPIOCounter + 1
ElseIf ((TempResult And &H4) <> &H4) Then
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2: Image Fail"
    'TempImageCounter = TempImageCounter + 1
    
ElseIf (TempResult And &H20) <> &H20 Then
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2: SRAM Fail"
    'TempSRAMCounter = TempSRAMCounter + 1
    
ElseIf (TempResult And &HC0) <> &H40 Then
    TestResult = "Bin4"
    MPTester.TestResultLab = "Bin4: CSet Fail"
    'TempCSETCounter = TempCSETCounter + 1

ElseIf (TempResult And &H18) = &H0 Then
    TestResult = "Bin3"
    MPTester.TestResultLab = "Bin3: LDO Fail"
    'TempLDOCounter = TempLDOCounter + 1
    
ElseIf (TempResult And &H18) = &H8 Then
    TestResult = "Bin5"
    MPTester.TestResultLab = "Bin5: Secondary PASS"
    'TempConditionCounter = TempConditionCounter + 1
ElseIf (TempResult = &H7F) Then
    TestResult = "PASS"
    MPTester.TestResultLab = "Bin1: PASS"
    'TempPASSCounter = TempPASSCounter + 1
Else
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2: Undefine Fail"
End If
                            
If TestResult = "Bin2" And FailCloseAP Then
    Call CloseVedioCap
End If
                            
'TempCounter = TempCounter + 1
'If TempCounter = 30 Then
'    'debug.print "PASS: " & TempPASSCounter & " ;SRAM: " & TempSRAMCounter & " ;CSET: "; TempCSETCounter _
'                ; " ;GPIO: " & TempGPIOCounter & " ;Condition: " & TempConditionCounter & " ;Image: " & TempImageCounter _
'                ; " ;LDO: " & TempLDOCounter & " ;VDD18: " & TempVDD18Counter
'
'    TempCounter = 0
'    TempSRAMCounter = 0
'    TempCSETCounter = 0
'    TempGPIOCounter = 0
'    TempConditionCounter = 0
'    TempPASSCounter = 0
'    TempImageCounter = 0
'    TempLDOCounter = 0
'    TempVDD18Counter = 0
'
'End If
                            
End Sub
Public Sub AU3825A61AFS7ETestSub()
'add unload driver function

'2012/8/30: base on ST1, skip LDO test & enhance SRAM test 7 cycle

 If PCI7248InitFinish = 0 Then
       Call PCI7248Exist
 End If
 
 If Not SetP1CInput_Flag Then
    result = DIO_PortConfig(card, Channel_P1C, INPUT_PORT)
    If result <> 0 Then
        MsgBox " config PCI_P1C as input card fail"
        End
    End If
    SetP1CInput_Flag = True
 End If
 
 Dim i As Integer
 Dim OldTimer As Long
 Dim PassTime As Long
 Dim rt2 As Long
 Dim LDOValue As Long
 Dim mMsg As MSG
 Dim LDORetry As Byte
 Dim TmpStr As String
 Dim TempResult As Byte
 Dim GPIO_Value As Long
 Dim TempCount As Byte
 Dim SRAMPASSCount As Byte
 Dim V18FailCount As Integer
 Dim SecondaryCount As Integer
 Dim LDOPASSCount As Integer
 ReDim LDOVal(1 To 50) As Long
 
 TestResult = ""
 TempResult = 0
 LDORetry = 0
 AlcorMPMessage = 0
 FW_Fail_Flag = False
 
cardresult = DO_WritePort(card, Channel_P1A, &HFE) 'Open ENA Power 1111_1110

Call MsecDelay(0.2)
If Not WaitDevOn("vid_058f") Then
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2:Vid/Pid UnKnow Fail"
    cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
    Call CloseVedioCap
    Exit Sub
End If
 
Call MsecDelay(0.8)
 
MPTester.TestResultLab = ""
'===============================================================
' Fail location initial
'===============================================================

If OldChipName <> ChipName Then
    
    ChDir App.Path & "\CamTest\AU3825A61FTTest_28QFN\"
    
    If Dir("C:\WINDOWS\system32\drivers\allow.sys") = "allow.sys" Then
        Kill ("C:\WINDOWS\system32\drivers\allow.sys")
    End If
    FWFail_Counter = 0
    Call CloseVedioCap
    OldChipName = ChipName
    
    Call Load_VerifyFW_Tool_AU3825_28QFN
    
    KillProcess ("VerifyFW_v3.2.4.exe")
    
End If

If FindWindow(vbNullString, "VideoCap") = 0 Then
    
    MPTester.Print "wait for VideoCap Ready"
    
    OldTimer = Timer
    
    If LoadVedioCap_AU3825_28QFN Then
        MPTester.Print "Ready Time="; Timer - OldTimer
    Else
        MPTester.TestResultLab = "Bin2:VideoCap Ready Fail "
        TestResult = "Bin2"
        MPTester.Print "VideoCap Ready Fail"
        cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
        Exit Sub
   End If
End If
         

'TempResult
'Bit 1,2  : GPIO_Setting        11: PASS
'Bit 3    : Image pattern        1: PASS
'Bit 4,5  : LDO                 11: PASS, 01:Condition PASS
'Bit 6    : SRAM                 1: PASS
'Bit 7,8    : CSET Value        01: PASS

'======================================
'   CSet Value Test
'======================================
MPTester.Print "Begin CSet Test........"
Call MsecDelay(0.04)
TempResult = TempResult + (CSET_Value_Test * &H40)

If (TempResult And &HC0) = &H40 Then
    MPTester.Print "CSet Value: PASS"
End If
Call MsecDelay(0.02)


'=====================================
'   GPIO Setting
'=====================================
Call MsecDelay(0.1)
MPTester.Print "Begin GPIO Setting Test........"

TempCount = 0

Do
    Call GPIO_Setting(&H20, &H54)
    Call MsecDelay(0.2)
    cardresult = DO_ReadPort(card, Channel_P1C, GPIO_Value)
    Call MsecDelay(0.02)

    If GPIO_Value = &HF9 Then
        TempResult = TempResult + 1
        Exit Do
    End If
    
    TempCount = TempCount + 1
    Call MsecDelay(0.02)
    
Loop Until (TempCount > 10)


TempCount = 0

Do
    Call GPIO_Setting(&H20, &H29)
    Call MsecDelay(0.2)
    cardresult = DO_ReadPort(card, Channel_P1C, GPIO_Value)
    Call MsecDelay(0.02)
    
    If GPIO_Value = &HDE Then
        TempResult = TempResult + 2
        Exit Do
    End If
    
    TempCount = TempCount + 1
    Call MsecDelay(0.02)
Loop Until (TempCount > 10)


If (TempResult And &H3) = 3 Then
    MPTester.Print "GPIO Setting: PASS"
Else
    GoTo TestEnd
End If


'=====================================
'   Image
'=====================================
Call MsecDelay(0.2)
MPTester.Print "Begin Image Test........"
'OldTimer = Timer

TempResult = TempResult + (Image_Test * &H4)

If (TempResult And &H4) = &H4 Then
    MPTester.Print "Image Test: PASS"
Else
    GoTo TestEnd
End If

'======================================
'   Set LV & SRAM Test
'======================================

cardresult = DO_WritePort(card, Channel_P1A, &HFC) 'Open ENA Power 1111_1100
Call MsecDelay(0.8)

For SRAMPASSCount = 1 To 7
    If SRAM_Test = 0 Then
        MPTester.Print "SRAM Test " & " Cycle " & SRAMPASSCount & ": Fail"
        Exit For
    Else
        MPTester.Print "Cycle " & SRAMPASSCount & ": PASS"
    End If
    
    If SRAMPASSCount = 7 Then
        TempResult = TempResult + (SRAM_Test * &H20)

        If (TempResult And &H20) = &H20 Then
            MPTester.Print "SRAM Test: PASS"
        End If
    End If
Next


TestEnd:

KillProcess ("VerifyFW_v3.2.4.exe")

cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
WaitDevOFF ("058f")
WaitDevOFF ("058f")
Call MsecDelay(0.2)

If FW_Fail_Flag Then
    FWFail_Counter = FWFail_Counter + 1
Else
    FWFail_Counter = 0
End If

If FWFail_Counter > 10 Then
    MPTester.FWFail_Label.Visible = True
Else
    MPTester.FWFail_Label.Visible = False
End If


If (TempResult And &H3) <> &H3 Then
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2: GPIO Setting Fail"
    'TempGPIOCounter = TempGPIOCounter + 1
ElseIf ((TempResult And &H4) <> &H4) Then
    TestResult = "Bin3"
    MPTester.TestResultLab = "Bin3: Image Fail"
    'TempImageCounter = TempImageCounter + 1

ElseIf (TempResult And &H20) <> &H20 Then
    TestResult = "Bin5"
    MPTester.TestResultLab = "Bin5: SRAM Fail"
    'TempSRAMCounter = TempSRAMCounter + 1
    
ElseIf (TempResult And &HC0) <> &H40 Then
    TestResult = "Bin4"
    MPTester.TestResultLab = "Bin4: CSet Fail"
    'TempCSETCounter = TempCSETCounter + 1

ElseIf (TempResult = &H67) Then
    TestResult = "PASS"
    MPTester.TestResultLab = "Bin1: PASS"
    'TempPASSCounter = TempPASSCounter + 1
Else
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2: Undefine Fail"
End If
                            
If TestResult = "Bin2" And FailCloseAP Then
    Call CloseVedioCap
End If

                            
End Sub

Public Sub AU3825A61AFS6ETestSub()
'add unload driver function
'2012/8/27: This code copy from AU3825A61BFF2E
'           purpose to sorting bin2 all fail binning
'
'           Bin2: Unknow
'           Bin3: GPIO、Image
'           Bin4: VDD18
'           Bin5: SRAM

 If PCI7248InitFinish = 0 Then
       Call PCI7248Exist
 End If
 
 If Not SetP1CInput_Flag Then
    result = DIO_PortConfig(card, Channel_P1C, INPUT_PORT)
    If result <> 0 Then
        MsgBox " config PCI_P1C as input card fail"
        End
    End If
    SetP1CInput_Flag = True
 End If
 
 Dim i As Integer
 Dim OldTimer As Long
 Dim PassTime As Long
 Dim rt2 As Long
 Dim LDOValue As Long
 Dim mMsg As MSG
 Dim LDORetry As Byte
 Dim TmpStr As String
 Dim TempResult As Byte
 Dim GPIO_Value As Long
 Dim TempCount As Byte
 Dim SRAMPASSCount As Byte
 Dim V18FailCount As Integer
 Dim SecondaryCount As Integer
 Dim LDOPASSCount As Integer
 ReDim LDOVal(1 To 50) As Long
 
 TestResult = ""
 TempResult = 0
 LDORetry = 0
 AlcorMPMessage = 0
 FW_Fail_Flag = False
 
cardresult = DO_WritePort(card, Channel_P1A, &HFE) 'Open ENA Power 1111_1110

Call MsecDelay(0.2)
If Not WaitDevOn("vid_058f") Then
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2:Vid/Pid UnKnow Fail"
    cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
    Call CloseVedioCap
    Exit Sub
End If
 
Call MsecDelay(0.8)
 
MPTester.TestResultLab = ""
'===============================================================
' Fail location initial
'===============================================================

If OldChipName <> ChipName Then
    
    ChDir App.Path & "\CamTest\AU3825A61FTTest_28QFN\"
    
    If Dir("C:\WINDOWS\system32\drivers\allow.sys") = "allow.sys" Then
        Kill ("C:\WINDOWS\system32\drivers\allow.sys")
    End If
    FWFail_Counter = 0
    Call CloseVedioCap
    OldChipName = ChipName
    
    Call Load_VerifyFW_Tool_AU3825_28QFN
    
    KillProcess ("VerifyFW_v3.2.4.exe")
    
End If

If FindWindow(vbNullString, "VideoCap") = 0 Then
    
    MPTester.Print "wait for VideoCap Ready"
    
    OldTimer = Timer
    
    If LoadVedioCap_AU3825_28QFN Then
        MPTester.Print "Ready Time="; Timer - OldTimer
    Else
        MPTester.TestResultLab = "Bin2:VideoCap Ready Fail "
        TestResult = "Bin2"
        MPTester.Print "VideoCap Ready Fail"
        cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
        Exit Sub
   End If
End If
         

'TempResult
'Bit 1,2  : GPIO_Setting        11: PASS
'Bit 3    : Image pattern        1: PASS
'Bit 4,5  : LDO                 11: PASS, 01:Condition PASS
'Bit 6    : SRAM                 1: PASS
'Bit 7,8    : CSET Value        01: PASS


'=====================================
'   GPIO Setting
'=====================================
Call MsecDelay(0.1)
MPTester.Print "Begin GPIO Setting Test........"

TempCount = 0

Do
    Call GPIO_Setting(&H20, &H54)
    Call MsecDelay(0.2)
    cardresult = DO_ReadPort(card, Channel_P1C, GPIO_Value)
    Call MsecDelay(0.02)

    If GPIO_Value = &HF9 Then
        TempResult = TempResult + 1
        Exit Do
    End If
    
    TempCount = TempCount + 1
    Call MsecDelay(0.02)
    
Loop Until (TempCount > 10)


TempCount = 0

Do
    Call GPIO_Setting(&H20, &H29)
    Call MsecDelay(0.2)
    cardresult = DO_ReadPort(card, Channel_P1C, GPIO_Value)
    Call MsecDelay(0.02)
    
    If GPIO_Value = &HDE Then
        TempResult = TempResult + 2
        Exit Do
    End If
    
    TempCount = TempCount + 1
    Call MsecDelay(0.02)
Loop Until (TempCount > 10)


If (TempResult And &H3) = 3 Then
    MPTester.Print "GPIO Setting: PASS"
Else
    GoTo TestEnd
End If


'=====================================
'   Image
'=====================================
Call MsecDelay(0.2)
MPTester.Print "Begin Image Test........"
'OldTimer = Timer

TempResult = TempResult + (Image_Test * &H4)

If (TempResult And &H4) = &H4 Then
    MPTester.Print "Image Test: PASS"
Else
    GoTo TestEnd
End If


'=====================================
'   LDO
'=====================================
MPTester.Print "Begin LDO Test........"

For i = 1 To 50
    cardresult = DO_ReadPort(card, Channel_P1B, LDOVal(i))
    Call MsecDelay(0.01)
Next

For i = 1 To 50
    If (LDOVal(i) And &H3) <> 0 Then                            'VDD18 Fail
        V18FailCount = V18FailCount + 1
    ElseIf (LDOVal(i) = &H40) Or (LDOVal(i) = &H20) Then    'XSA < 2.65 or XSA > 2.95
        SecondaryCount = SecondaryCount + 1
    ElseIf LDOVal(i) = 0 Then
        LDOPASSCount = LDOPASSCount + 1
    End If
Next

'MPTester.Print "V18: " & V18FailCount
'MPTester.Print "Condition: " & SecondaryCount
'MPTester.Print "LDO PSS: " & LDOPASSCount

If V18FailCount >= 1 Then
    MPTester.Print "Bin4: VDD18 Fail"
    TestResult = "Bin4"
    MPTester.TestResultLab = "Bin4: VDD18 Fail"
    cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'Close ENA Power 1111_1111
    'TempCounter = TempCounter + 1
    'TempVDD18Counter = TempVDD18Counter + 1
    Call CloseVedioCap
    Exit Sub
End If

'If LDOValue = 0 Then
'    MPTester.Print "LDO PASS"
'    TempResult = TempResult + &H18
'ElseIf (LDOValue = &H40) Or (LDOValue = &H20) Then      'XSA < 2.65 or XSA > 2.95
'    MPTester.Print "LDO Condition PASS"
'    TempResult = TempResult + &H8
'ElseIf (LDOValue And &H3) <> 0 Then                     'VDD18 Fail
'    MPTester.Print "Bin2: VDD18 Fail"
'    TestResult = "Bin2"
'    MPTester.TestResultLab = "Bin2: VDD18 Fail"
'    cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'Close ENA Power 1111_1111
'    TempCounter = TempCounter + 1
'    TempVDD18Counter = TempVDD18Counter + 1
'    Exit Sub
'Else
'    MPTester.Print "LDO Fail"
'End If


'======================================
'   Set LV & SRAM Test
'======================================

cardresult = DO_WritePort(card, Channel_P1A, &HFC) 'Open ENA Power 1111_1100
Call MsecDelay(0.8)

For SRAMPASSCount = 1 To 2
    If SRAM_Test = 0 Then
        MPTester.Print "SRAM Test " & " Cycle " & SRAMPASSCount & ": Fail"
        Exit For
    Else
        MPTester.Print "Cycle " & SRAMPASSCount & ": PASS"
    End If
    
    If SRAMPASSCount = 2 Then
        TempResult = TempResult + (SRAM_Test * &H20)

        If (TempResult And &H20) = &H20 Then
            MPTester.Print "SRAM Test: PASS"
        End If
    End If
Next


TestEnd:

KillProcess ("VerifyFW_v3.2.4.exe")

cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
WaitDevOFF ("058f")
WaitDevOFF ("058f")
Call MsecDelay(0.2)

If FW_Fail_Flag Then
    FWFail_Counter = FWFail_Counter + 1
Else
    FWFail_Counter = 0
End If

If FWFail_Counter > 10 Then
    MPTester.FWFail_Label.Visible = True
Else
    MPTester.FWFail_Label.Visible = False
End If


If (TempResult And &H3) <> &H3 Then
    TestResult = "Bin3"
    MPTester.TestResultLab = "Bin3: GPIO Setting Fail"
    'TempGPIOCounter = TempGPIOCounter + 1
ElseIf ((TempResult And &H4) <> &H4) Then
    TestResult = "Bin3"
    MPTester.TestResultLab = "Bin3: Image Fail"
    'TempImageCounter = TempImageCounter + 1
    
ElseIf (TempResult And &H20) <> &H20 Then
    TestResult = "Bin5"
    MPTester.TestResultLab = "Bin5: SRAM Fail"
    'TempSRAMCounter = TempSRAMCounter + 1

'ElseIf (TempResult And &HC0) <> &H40 Then
'    TestResult = "Bin4"
'    MPTester.TestResultLab = "Bin4: CSet Fail"
    'TempCSETCounter = TempCSETCounter + 1

'ElseIf (TempResult And &H18) = &H0 Then
'    TestResult = "Bin3"
'    MPTester.TestResultLab = "Bin3: LDO Fail"
    'TempLDOCounter = TempLDOCounter + 1
    
'ElseIf (TempResult And &H18) = &H8 Then
'    TestResult = "Bin5"
'    MPTester.TestResultLab = "Bin5: Secondary PASS"
    'TempConditionCounter = TempConditionCounter + 1
'ElseIf (TempResult = &H7F) Then
'    TestResult = "PASS"
'    MPTester.TestResultLab = "Bin1: PASS"
    'TempPASSCounter = TempPASSCounter + 1
Else
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2: Undefine Fail"
End If
                            
If TestResult = "Bin2" And FailCloseAP Then
    Call CloseVedioCap
End If
                            
'TempCounter = TempCounter + 1
'If TempCounter = 30 Then
'    'debug.print "PASS: " & TempPASSCounter & " ;SRAM: " & TempSRAMCounter & " ;CSET: "; TempCSETCounter _
'                ; " ;GPIO: " & TempGPIOCounter & " ;Condition: " & TempConditionCounter & " ;Image: " & TempImageCounter _
'                ; " ;LDO: " & TempLDOCounter & " ;VDD18: " & TempVDD18Counter
'
'    TempCounter = 0
'    TempSRAMCounter = 0
'    TempCSETCounter = 0
'    TempGPIOCounter = 0
'    TempConditionCounter = 0
'    TempPASSCounter = 0
'    TempImageCounter = 0
'    TempLDOCounter = 0
'    TempVDD18Counter = 0
'
'End If
                            
End Sub
Public Sub AU3825A61AFQ2ETestSub()

'2012/8/13 EQC fail lot sorting program

 If PCI7248InitFinish = 0 Then
       Call PCI7248Exist
 End If
 
 If Not SetP1CInput_Flag Then
    result = DIO_PortConfig(card, Channel_P1C, INPUT_PORT)
    If result <> 0 Then
        MsgBox " config PCI_P1C as input card fail"
        End
    End If
    SetP1CInput_Flag = True
 End If
 
 Dim i As Integer
 Dim OldTimer As Long
 Dim PassTime As Long
 Dim rt2 As Long
 Dim LDOValue As Long
 Dim mMsg As MSG
 Dim LDORetry As Byte
 Dim TmpStr As String
 Dim TempResult As Byte
 Dim GPIO_Value As Long
 Dim TempCount As Byte
 Dim SRAMPASSCount As Byte
 Dim V18FailCount As Integer
 Dim SecondaryCount As Integer
 Dim LDOPASSCount As Integer
 ReDim LDOVal(1 To 50) As Long
 
 TestResult = ""
 TempResult = 0
 LDORetry = 0
 AlcorMPMessage = 0
 FW_Fail_Flag = False
 
cardresult = DO_WritePort(card, Channel_P1A, &HFE) 'Open ENA Power 1111_1110

Call MsecDelay(0.2)
If Not WaitDevOn("vid_058f") Then
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2:Vid/Pid UnKnow Fail"
    cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
    Call CloseVedioCap
    Exit Sub
End If
 
Call MsecDelay(0.8)
 
MPTester.TestResultLab = ""
'===============================================================
' Fail location initial
'===============================================================

If OldChipName <> ChipName Then
    
    ChDir App.Path & "\CamTest\AU3825A61FTTest_28QFN\"
    
    If Dir("C:\WINDOWS\system32\drivers\allow.sys") = "allow.sys" Then
        Kill ("C:\WINDOWS\system32\drivers\allow.sys")
    End If
    FWFail_Counter = 0
    Call CloseVedioCap
    OldChipName = ChipName
    
    Call Load_VerifyFW_Tool_AU3825_28QFN
    
    KillProcess ("VerifyFW_v3.2.4.exe")
    
End If

If FindWindow(vbNullString, "VideoCap") = 0 Then
    
    MPTester.Print "wait for VideoCap Ready"
    
    OldTimer = Timer
    
    If LoadVedioCap_AU3825_28QFN Then
        MPTester.Print "Ready Time="; Timer - OldTimer
    Else
        MPTester.TestResultLab = "Bin2:VideoCap Ready Fail "
        TestResult = "Bin2"
        MPTester.Print "VideoCap Ready Fail"
        cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
        Exit Sub
   End If
End If
         

'TempResult
'Bit 1,2  : GPIO_Setting        11: PASS
'Bit 3    : Image pattern        1: PASS
'Bit 4,5  : LDO                 11: PASS, 01:Condition PASS
'Bit 6    : SRAM                 1: PASS
'Bit 7,8    : CSET Value        01: PASS

'======================================
'   Set LV & SRAM Test
'======================================

cardresult = DO_WritePort(card, Channel_P1A, &HFC) 'Open ENA Power 1111_1100
Call MsecDelay(0.8)

For SRAMPASSCount = 1 To 7
    If SRAM_Test = 0 Then
        MPTester.Print "SRAM Test " & " Cycle " & SRAMPASSCount & ": Fail"
        Exit For
    Else
        MPTester.Print "Cycle " & SRAMPASSCount & ": PASS"
    End If
    
    If SRAMPASSCount = 7 Then
        TempResult = TempResult + (SRAM_Test * &H20)

        If (TempResult And &H20) = &H20 Then
            MPTester.Print "SRAM Test: PASS"
        End If
    End If
Next


TestEnd:

KillProcess ("VerifyFW_v3.2.4.exe")

cardresult = DO_WritePort(card, Channel_P1A, &HFF) 'Close ENA Power 1111_1111
WaitDevOFF ("058f")
WaitDevOFF ("058f")
Call MsecDelay(0.2)


If (TempResult And &H20) <> &H20 Then
    TestResult = "Bin3"
    MPTester.TestResultLab = "Bin3: SRAM Fail"
ElseIf (TempResult And &H20) = &H20 Then
    TestResult = "PASS"
    MPTester.TestResultLab = "Bin1: PASS"
Else
    TestResult = "Bin2"
    MPTester.TestResultLab = "Bin2: Undefine Fail"
End If
                            
If TestResult = "Bin2" And FailCloseAP Then
    Call CloseVedioCap
End If
                            
End Sub

Public Sub CloseVedioCap()
Dim rt2 As Long
Dim TempOldTime As Long
Dim TempPassingTime As Long
Dim mMsg As MSG

    winHwnd = FindWindow(vbNullString, "VideoCap")
    
    If winHwnd = 0 Then
        Exit Sub
    End If
    
    TempOldTime = Timer
    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
         
        TempPassingTime = Timer - TempOldTime
       
    Loop Until AlcorMPMessage = WM_CAM_MP_READY _
          Or TempPassingTime > 3
    
    TempOldTime = Timer
    If winHwnd <> 0 Then
        Do
            rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
            Call MsecDelay(0.5)
            winHwnd = FindWindow(vbNullString, "VideoCap")
            TempPassingTime = Timer - TempOldTime
        Loop While winHwnd <> 0 _
            And TempPassingTime < 2
    End If
    
    KillProcess ("VideoCap.exe")
    
End Sub

Public Sub Load_AU3825_40QFN_FW_Update()

Dim OldPath As String
Dim OldTime As Long

    OldPath = CurDir
    ChDir (App.Path & "\CamTest\AU3825A61FTTest_40QFN\FW update Tool\")

    
    winHwnd = FindWindow(vbNullString, "AlcorMPTool V3.12.620")
    
    OldTime = Timer
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\CamTest\AU3825A61FTTest_40QFN\FW update Tool\MPTool_lite_v3.12.620.exe", "", "", SW_SHOW)
    End If
    
    SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags
    
    Call MsecDelay(2#)
    
    Do
        winHwnd = FindWindow(vbNullString, "AlcorMPTool V3.12.620")
        MsecDelay (0.5)
    Loop While (winHwnd <> 0) And (Timer - OldTime < 15)
    
    
    FW_Fail_Flag = False
    
    ChDir OldPath
    
End Sub

Public Sub Load_AU3825_28QFN_FW_Update()

Dim OldPath As String
Dim OldTime As Long

    OldPath = CurDir
    ChDir (App.Path & "\CamTest\AU3825A61FTTest_28QFN\FW update Tool\")

    
    winHwnd = FindWindow(vbNullString, "AlcorMPTool V3.12.620")
    
    OldTime = Timer
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\CamTest\AU3825A61FTTest_28QFN\FW update Tool\MPTool_lite_v3.12.620.exe", "", "", SW_SHOW)
    End If
    
    SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags
    
    Call MsecDelay(2#)
    
    Do
        winHwnd = FindWindow(vbNullString, "AlcorMPTool V3.12.620")
        MsecDelay (0.5)
    Loop While (winHwnd <> 0) And (Timer - OldTime < 15)
    
    
    FW_Fail_Flag = False
    
    ChDir OldPath
    
End Sub
