Attribute VB_Name = "AU2100Series"
Public Const WM_USER = &H400
Public Const WM_FT_TEST_READY = WM_USER + &H800
Public Const WM_MP_TEST_READY = WM_USER + &H850

Public Const WM_IsConnected_CK2 = WM_USER + &H700
Public Const WM_Connected_CK2 = WM_USER + &H710
Public Const WM_DisConnected_CK2 = WM_USER + &H720
Public Const WM_ReStart_CK2 = WM_USER + &H750
Public Const WM_PWROff_Dev = WM_USER + &H760    'Power off CK2 GPIO2

Public Const WM_SetPinNum_16 = WM_USER + &H100
Public Const WM_SetPinNum_24 = WM_USER + &H110
Public Const WM_SetPinNum_32 = WM_USER + &H120

Public Const WM_SetInput_0 = WM_USER + &H200
Public Const WM_SetInput_1 = WM_USER + &H210
Public Const WM_SetRead = WM_USER + &H220
Public Const WM_SetOutput_0 = WM_USER + &H230
Public Const WM_SetOutput_1 = WM_USER + &H240

Public Const WM_ClearList = WM_USER + &H300
Public Const WM_Connect_CK2 = WM_USER + &H350
Public Const WM_DisConnect_CK2 = WM_USER + &H360

Public Const WM_SetSuccess = WM_USER + &H400
Public Const WM_ReadPASS = WM_USER + &H500
Public Const WM_SetFail = WM_USER + &H600

Public Const WM_MP_START = WM_USER + &H700
Public Const WM_MP_FAIL = WM_USER + &H710
Public Const WM_MP_PASS = WM_USER + &H720
Public Const WM_MP_Connect_CK2 = WM_USER + &H900
Public Const WM_MP_Connect_PASS = WM_USER + &H910
Public Const WM_MP_Connect_FAIL = WM_USER + &H920
Public Const WM_MP_DisConnect_CK2 = WM_USER + &H930
Public Const WM_MP_DisConnect_PASS = WM_USER + &H940
Public Const WM_MP_DisConnect_FAIL = WM_USER + &H950

Public Const WM_MP_ReadFWFromDevice = WM_USER + &H600
Public Const WM_MP_ReadFW_CusID = WM_USER + &H610
Public Const WM_MP_ReadFW_DevID = WM_USER + &H620
Public Const WM_MP_ReadFW_ResNum = WM_USER + &H630
Public Const WM_MP_ReadFW_ExtFW = WM_USER + &H640
Public Const WM_MP_ReadFW_SrcFW = WM_USER + &H650
Public Const WM_MP_ReadFW_FWLen = WM_USER + &H660
Public Const WM_MP_ReadFW_CrcVal = WM_USER + &H670
Public Const WM_MP_ReadFW_Ret = WM_USER + &H690


Public Const AU2100_MPTool_Title = "AmMaserati_MPTool"
Public Const AU2100_AP_Title = "AU2100_Test_Tool"
Public Const AU2100_NewMPTool_Title = "Program Tool"

Public Const AU2101_MPTool_Title = "AmMaseratiTool Version 20140521_RD"
Public Const AU2101_AP_Title = "AU2101_Test_Tool"
Public Const AU2101_FPTool_Title = "AmMaseratiTool Version 20140717_Release"

Public bAU2100_OTP As Boolean
Public OpenArg As String
Public AU2100_ContiFail As Integer
Public AU2100_MP_CMD As String
Public Target_CusID As Long
Public Target_DevID As Long
Public Target_ResNum As Long
Public Target_ExtFW As Long
Public Target_SrcFW As Long
Public Target_FWLen As Long
Public Target_CrcVal As Long
Public AU2101_DllName As String
Public AU2101_BinName As String

'Public AU2101_MP_CMD As String
'Public AU2100_MP_ARG2_ini As String

Public Sub AU2100_NameSub()
    
    AU2100_MP_CMD = "Touch_FLASH.Bin TouchPanel.ini"
    'AU2100_MP_ARG2_ini = ""
    
    If ChipName = "AU2100A41DFF20" Then
        bAU2100_OTP = False
        OpenArg = "32"
    ElseIf ChipName = "AU2100A41DFM20" Then
        bAU2100_OTP = True
        OpenArg = "32"
    ElseIf ChipName = "AU2100A41BFF20" Then
        bAU2100_OTP = False
        OpenArg = "24"
    ElseIf ChipName = "AU2100A41BFM20" Then
        bAU2100_OTP = True
        OpenArg = "24"
    ElseIf ChipName = "AU2100A41CFF20" Then
        bAU2100_OTP = False
        OpenArg = "16"
    ElseIf ChipName = "AU2100A41CFM20" Then
        bAU2100_OTP = True
        OpenArg = "16"
    End If
    
    Call AU2100TestSub

End Sub

Public Sub AU2101_NameSub()
    
    If ChipName = "AU2101A41AFF20" Then
        bAU2100_OTP = False
        OpenArg = "32"
    ElseIf ChipName = "AU2101A41AFM20" Then
        bAU2100_OTP = True
        OpenArg = "32"
    ElseIf ChipName = "AU2101A41BFF20" Then
        bAU2100_OTP = False
        OpenArg = "24"
    ElseIf ChipName = "AU2101A41BFM20" Then
        bAU2100_OTP = True
        OpenArg = "24"
    ElseIf ChipName = "AU2101A41CFF20" Then
        bAU2100_OTP = False
        OpenArg = "16"
    ElseIf ChipName = "AU2101A41CFM20" Then
        bAU2100_OTP = True
        OpenArg = "16"
    ElseIf ChipName = "AU2101B41DFF20" Then     ' A41GCF = B41BDF改腳位
        bAU2100_OTP = False
        OpenArg = "16"
    ElseIf ChipName = "AU2101B41DFM20" Then     ' A41GCF = B41BDF改腳位
        bAU2100_OTP = True
        OpenArg = "16"
    End If
    
    Call AU2101TestSub

End Sub

Public Sub AU2101_FPXXXX_NameSub()

    If ChipName = "AU2101DFFP0002" Then
        bAU2100_OTP = True
        OpenArg = "16"
        Target_CusID = 2
        Target_DevID = 2
        Target_ResNum = &H780C5044
        Target_ExtFW = 0
        Target_SrcFW = &HAAA30715
        Target_FWLen = 62
        Target_CrcVal = &H3989
        AU2101_DllName = "TouchPanel_Test_0729.dll"
        AU2101_BinName = "2014072900_0002_0002.bin"
    ElseIf ChipName = "AU2101HFFP1403" Then
        bAU2100_OTP = True
        OpenArg = "16"
        Target_CusID = 2
        Target_DevID = 2
        Target_ResNum = &H780C9A7C
        Target_ExtFW = 0
        Target_SrcFW = &HAAA30814
        'Target_SrcFW = -1432156140
        Target_FWLen = 62
        Target_CrcVal = 65017
        AU2101_DllName = "TouchPanel_Test_20140919.dll"
        AU2101_BinName = "2014091900_0002_0002.bin"
    ElseIf ChipName = "AU2101HFFP1404" Then
        bAU2100_OTP = True
        OpenArg = "16"
        Target_CusID = 2
        Target_DevID = 2
        Target_ResNum = &H780C9EC8
        Target_ExtFW = 0
        Target_SrcFW = &HAAA30814
        Target_FWLen = 62
        Target_CrcVal = 35076
        AU2101_DllName = "TouchPanel_Test_20140930.dll"
        AU2101_BinName = "Touch_FLASH_20140930.Bin"
    ElseIf ChipName = "AU2101HFFP1501" Then
        bAU2100_OTP = True
        OpenArg = "16"
        Target_CusID = 2
        Target_DevID = 2
        Target_ResNum = &H781AEC80
        Target_ExtFW = 0
        Target_SrcFW = &HAAA30814
        Target_FWLen = 62
        Target_CrcVal = 62122
        AU2101_DllName = "TouchPanel_Test_20150304.dll"
        AU2101_BinName = "Touch_FLASH_20150304.Bin"
    End If
    
    Call AU2101_FPXXXX_Sub
    
End Sub


Public Sub AU2100_ProgNameSub()

    If ChipName = "AU2100CFFP0101" Then
        bAU2100_OTP = False
        AU2100_MP_CMD = App.Path & "\TouchKey\IntegrateIniVer\1302_2013091701.bin"
    End If
    
    Call AU2100ProgSub
    
End Sub

Public Sub LoadAP_Click_AU2100(Arg As String)

Dim TimePass
Dim rt2

    ' find window
    winHwnd = FindWindow(vbNullString, AU2100_AP_Title)
 
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\TouchKey\" & "AU2100_Test_Tool.exe", Arg, "", SW_SHOW)
    End If
    SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags

End Sub

Public Sub LoadAP_Click_AU2101(Arg As String)

Dim TimePass
Dim rt2

    ' find window
    winHwnd = FindWindow(vbNullString, AU2101_AP_Title)
 
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\TouchKey\AU2101\" & "AU2101_Test_Tool.exe", Arg, "", SW_SHOW)
    End If
    SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags

End Sub

Public Sub LoadMP_Click_AU2100()

Dim TimePass
Dim rt2

    ' find window
    winHwnd = FindWindow(vbNullString, AU2100_MPTool_Title)
 
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\TouchKey\" & "AmMaserati_MPTool.exe", AU2100_MP_CMD, "", SW_SHOW)
    End If
    SetWindowPos winHwnd, HWND_TOPMOST, 900, 300, 0, 0, Flags

End Sub

Public Sub LoadMP_Click_AU2101()

Dim TimePass
Dim rt2

    ' find window
    winHwnd = FindWindow(vbNullString, AU2101_MPTool_Title)
 
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\TouchKey\AU2101\" & "AmMaseratiTool.exe", "", "", SW_SHOW)
    End If
    SetWindowPos winHwnd, HWND_TOPMOST, 900, 300, 0, 0, Flags

End Sub

Public Sub LoadMP_Click_AU2101_FPXXXX(DllName As String, BinName As String)

Dim TimePass
Dim rt2

    ' find window
    winHwnd = FindWindow(vbNullString, AU2101_FPTool_Title)
 
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\TouchKey\AU2101\AU2101_FPXXXX\" & "AmMaseratiTool_FPXXXX.exe", DllName & " " & BinName, "", SW_SHOW)
    End If
    SetWindowPos winHwnd, HWND_TOPMOST, 900, 300, 0, 0, Flags

End Sub

Public Sub LoadNewMP_Click_AU2100()

Dim TimePass
Dim rt2

    ' find window
    winHwnd = FindWindow(vbNullString, AU2100_MPTool_Title)
 
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\TouchKey\IntegrateIniVer\" & "Program_Tool.exe", AU2100_MP_CMD, "", SW_SHOW)
    End If
    SetWindowPos winHwnd, HWND_TOPMOST, 900, 300, 0, 0, Flags

End Sub

Public Function StartAP_Click_AU2100(TestType As Long, TargetH_L As Long) As Boolean

Dim rt2
Dim OldTimer As Long
Dim PassTime As Long
Dim mMsg As MSG
   
    winHwnd = FindWindow(vbNullString, AU2100_AP_Title)
    rt2 = PostMessage(winHwnd, TestType, ByVal (TargetH_L), 0&)
    
    If (TestType = WM_DisConnect_CK2) Then
        Exit Function
    End If
    
    
    OldTimer = Timer
    AlcorMPMessage = 0
            
    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
        
        If (TestType = WM_SetRead) Then
            If (AlcorMPMessage = WM_SetSuccess) Then
                AlcorMPMessage = 0
            End If
        End If
        
        PassTime = Timer - OldTimer
    Loop Until AlcorMPMessage = WM_SetSuccess Or _
               AlcorMPMessage = WM_SetFail Or _
               AlcorMPMessage = WM_ReadPASS Or _
               PassTime > 5 Or _
               AlcorMPMessage = WM_CLOSE Or _
               AlcorMPMessage = WM_DESTROY
                  
    If (AlcorMPMessage = WM_SetSuccess) Or (AlcorMPMessage = WM_ReadPASS) Then
        StartAP_Click_AU2100 = True
    Else
        StartAP_Click_AU2100 = False
    End If

End Function

Public Function StartAP_Click_AU2101(TestType As Long, TargetH_L As Long) As Boolean

Dim rt2
Dim OldTimer As Long
Dim PassTime As Long
Dim mMsg As MSG
   
    winHwnd = FindWindow(vbNullString, AU2101_AP_Title)
    rt2 = PostMessage(winHwnd, TestType, ByVal (TargetH_L), 0&)
    
    If (TestType = WM_DisConnect_CK2) Then
        Exit Function
    End If
    
    
    OldTimer = Timer
    AlcorMPMessage = 0
            
    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
        
        If (TestType = WM_SetRead) Then
            If (AlcorMPMessage = WM_SetSuccess) Then
                AlcorMPMessage = 0
            End If
        End If
        
        PassTime = Timer - OldTimer
    Loop Until AlcorMPMessage = WM_SetSuccess Or _
               AlcorMPMessage = WM_SetFail Or _
               AlcorMPMessage = WM_ReadPASS Or _
               PassTime > 5 Or _
               AlcorMPMessage = WM_CLOSE Or _
               AlcorMPMessage = WM_DESTROY
                  
    If (AlcorMPMessage = WM_SetSuccess) Or (AlcorMPMessage = WM_ReadPASS) Then
        StartAP_Click_AU2101 = True
    Else
        StartAP_Click_AU2101 = False
    End If

End Function

Public Function Inquiry_CK2_Connected() As Boolean

Dim rt2
Dim OldTimer As Long
Dim PassTime As Long
Dim mMsg As MSG
   
    winHwnd = FindWindow(vbNullString, AU2100_AP_Title)
    
    If (winHwnd = 0) Then
        Inquiry_CK2_Connected = False
        Exit Function
    End If
    
    rt2 = PostMessage(winHwnd, WM_IsConnected_CK2, 0&, 0&)
   
    OldTimer = Timer
    AlcorMPMessage = 0
            
    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
        
        PassTime = Timer - OldTimer
    Loop Until AlcorMPMessage = WM_Connected_CK2 Or _
               AlcorMPMessage = WM_DisConnected_CK2 Or _
               PassTime > 2 Or _
               AlcorMPMessage = WM_CLOSE Or _
               AlcorMPMessage = WM_DESTROY
                  
    If (AlcorMPMessage = WM_Connected_CK2) Then
        Inquiry_CK2_Connected = True
    Else
        Inquiry_CK2_Connected = False
    End If

End Function

Public Function Inquiry_CK2_Connected_AU2101() As Boolean

Dim rt2
Dim OldTimer As Long
Dim PassTime As Long
Dim mMsg As MSG
   
    winHwnd = FindWindow(vbNullString, AU2101_AP_Title)
    
    If (winHwnd = 0) Then
        Inquiry_CK2_Connected_AU2101 = False
        Exit Function
    End If
    
    rt2 = PostMessage(winHwnd, WM_IsConnected_CK2, 0&, 0&)
   
    OldTimer = Timer
    AlcorMPMessage = 0
            
    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
        
        PassTime = Timer - OldTimer
    Loop Until AlcorMPMessage = WM_Connected_CK2 Or _
               AlcorMPMessage = WM_DisConnected_CK2 Or _
               PassTime > 2 Or _
               AlcorMPMessage = WM_CLOSE Or _
               AlcorMPMessage = WM_DESTROY
                  
    If (AlcorMPMessage = WM_Connected_CK2) Then
        Inquiry_CK2_Connected_AU2101 = True
    Else
        Inquiry_CK2_Connected_AU2101 = False
    End If

End Function

Public Sub StartMP_Click_AU2100()

Dim rt2
   
    winHwnd = FindWindow(vbNullString, AU2100_MPTool_Title)
    If winHwnd <> 0 Then
        rt2 = PostMessage(winHwnd, WM_MP_START, 0&, 0&)
    End If
End Sub

Public Sub StartMP_Click_AU2101()

Dim rt2
   
    winHwnd = FindWindow(vbNullString, AU2101_MPTool_Title)
    If winHwnd <> 0 Then
        rt2 = PostMessage(winHwnd, WM_MP_START, 0&, 0&)
    End If
End Sub

Public Sub StartMP_Click_AU2101_FPXXXX()

Dim rt2
   
    winHwnd = FindWindow(vbNullString, AU2101_FPTool_Title)
    If winHwnd <> 0 Then
        rt2 = PostMessage(winHwnd, WM_MP_START, 0&, 0&)
    End If
End Sub

Public Sub MPConnectCK2_Click_AU2101()

Dim rt2
   
    winHwnd = FindWindow(vbNullString, AU2101_MPTool_Title)
    If winHwnd <> 0 Then
        rt2 = PostMessage(winHwnd, WM_MP_Connect_CK2, 0&, 0&)
    End If
End Sub

Public Sub MPConnectCK2_Click_AU2101_FPXXXX()

Dim rt2
   
    winHwnd = FindWindow(vbNullString, AU2101_FPTool_Title)
    If winHwnd <> 0 Then
        rt2 = PostMessage(winHwnd, WM_MP_Connect_CK2, 0&, 0&)
    End If
End Sub

Public Sub MPDisConnectCK2_Click_AU2101()

Dim rt2
   
    winHwnd = FindWindow(vbNullString, AU2101_MPTool_Title)
    If winHwnd <> 0 Then
        rt2 = PostMessage(winHwnd, WM_MP_DisConnect_CK2, 0&, 0&)
    End If
End Sub

Public Sub MPDisConnectCK2_Click_AU2101_FPXXXX()

Dim rt2
   
    winHwnd = FindWindow(vbNullString, AU2101_FPTool_Title)
    If winHwnd <> 0 Then
        rt2 = PostMessage(winHwnd, WM_MP_DisConnect_CK2, 0&, 0&)
    End If
End Sub

Public Sub MPReadFWVersion_Click_AU2101_FPXXXX()

Dim rt2
   
    winHwnd = FindWindow(vbNullString, AU2101_FPTool_Title)
    If winHwnd <> 0 Then
        rt2 = PostMessage(winHwnd, WM_MP_ReadFWFromDevice, 0&, 0&)
    End If
End Sub

Public Function MPReadFWVersion_Click_AU2101(CheckItem As Long) As Long

Dim rt2
Dim OldTimer As Long
Dim PassTime As Long
Dim mMsg As MSG
   
    winHwnd = FindWindow(vbNullString, AU2101_FPTool_Title)
    If winHwnd <> 0 Then
        rt2 = PostMessage(winHwnd, CheckItem, 0&, 0&)
    End If
    
    OldTimer = Timer
    AlcorMPMessage = 0
            
    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            MPReadFWVersion_Click_AU2101 = mMsg.wParam
            Debug.Print mMsg.wParam
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
        
        PassTime = Timer - OldTimer
    Loop Until AlcorMPMessage = WM_MP_ReadFW_Ret Or _
                                PassTime >= 2
                                
    If PassTime >= 2 Then
        MPReadFWVersion_Click_AU2101 = -1
        Exit Function
    End If
    
    

End Function

Public Sub StartNewMP_Click_AU2100()

Dim rt2
   
    winHwnd = FindWindow(vbNullString, AU2100_NewMPTool_Title)
    If winHwnd <> 0 Then
        rt2 = PostMessage(winHwnd, WM_MP_START, 0&, 0&)
    End If
End Sub

Public Sub Clear_Click_AU2100()

Dim rt2
   
    winHwnd = FindWindow(vbNullString, AU2100_AP_Title)
    If winHwnd <> 0 Then
        rt2 = PostMessage(winHwnd, WM_ClearList, 0&, 0&)
    End If
End Sub

Public Sub Clear_Click_AU2101()

Dim rt2
   
    winHwnd = FindWindow(vbNullString, AU2101_AP_Title)
    If winHwnd <> 0 Then
        rt2 = PostMessage(winHwnd, WM_ClearList, 0&, 0&)
    End If
End Sub

Public Sub ReStart_Click_AU2100()

Dim rt2
   
    winHwnd = FindWindow(vbNullString, AU2100_AP_Title)
    If winHwnd <> 0 Then
        rt2 = PostMessage(winHwnd, WM_ReStart_CK2, 0&, 0&)
    End If
End Sub

Public Sub ReStart_Click_AU2101()

Dim rt2
   
    winHwnd = FindWindow(vbNullString, AU2101_AP_Title)
    If winHwnd <> 0 Then
        rt2 = PostMessage(winHwnd, WM_ReStart_CK2, 0&, 0&)
    End If
End Sub

Public Sub SetPinNumber_AU2101()

Dim rt2
   
    winHwnd = FindWindow(vbNullString, AU2101_AP_Title)
    If winHwnd <> 0 Then
    
        If (OpenArg = "32") Then
            rt2 = PostMessage(winHwnd, WM_SetPinNum_32, 0&, 0&)
        ElseIf (OpenArg = "24") Then
            rt2 = PostMessage(winHwnd, WM_SetPinNum_24, 0&, 0&)
        ElseIf (OpenArg = "16") Then
            rt2 = PostMessage(winHwnd, WM_SetPinNum_16, 0&, 0&)
        End If
    End If
End Sub


Public Sub PWROff_Click_AU2100()

Dim rt2
   
    winHwnd = FindWindow(vbNullString, AU2100_AP_Title)
    If winHwnd <> 0 Then
        rt2 = PostMessage(winHwnd, WM_PWROff_Dev, 0&, 0&)
    End If
End Sub

Public Sub PWROff_Click_AU2101()

Dim rt2
   
    winHwnd = FindWindow(vbNullString, AU2101_AP_Title)
    If winHwnd <> 0 Then
        rt2 = PostMessage(winHwnd, WM_PWROff_Dev, 0&, 0&)
    End If
End Sub

Public Sub AU2100TestSub()
 
Dim OldTimer
Dim PassTime
Dim rt2
Dim mMsg As MSG
Dim TempRes As Byte
Dim GateVal As Long


    'P1A(Output)    P1B(Input)
    '8765 4321      8765 4321
    
    '   S SEEE      GGGG GGGG
    '   W WNNN      AAAA AAAA
    '   I I|||      TTTT TTTT
    '   T TCBA      EEEE EEEE
    '   C C
    '   H H         ---- OOOO
    '               AAAA RRRR
    '   E S         NNNN
    '   N E         DDDD
    '   A L
    
    'Switch ENA : 0: active
    'Switch Sel : 0: Device Input; 1: Device Output
    

    MPTester.TestResultLab = ""
    
    If PCI7248InitFinish = 0 Then
        PCI7248Exist
    End If
    
    result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
    AlcorMPMessage = 0
    NewChipFlag = 0
    
    
    If OldChipName <> ChipName Then
        ' reset program
        winHwnd = FindWindow(vbNullString, AU2100_AP_Title)
        If winHwnd <> 0 Then
            Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(1#)
                winHwnd = FindWindow(vbNullString, AU2100_AP_Title)
            Loop While winHwnd <> 0
        End If
        
        winHwnd = FindWindow(vbNullString, AU2100_MPTool_Title)
        If winHwnd <> 0 Then
            Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(1#)
                winHwnd = FindWindow(vbNullString, AU2100_MPTool_Title)
            Loop While winHwnd <> 0
        End If
        
        winHwnd = FindWindow(vbNullString, AU2100_NewMPTool_Title)
        If winHwnd <> 0 Then
            Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(1#)
                winHwnd = FindWindow(vbNullString, AU2100_MPTool_Title)
            Loop While winHwnd <> 0
        End If
        
        NewChipFlag = 1
    End If
              
    OldChipName = ChipName
    
    '===================================================================================================
    '============================================= MP Flow =============================================
    '===================================================================================================
    
    If (bAU2100_OTP) Then       'MP + FT2
    
        cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        
        If (Inquiry_CK2_Connected) Then
            Call StartAP_Click_AU2100(WM_DisConnect_CK2, 1)
        End If
        
        If (NewChipFlag = 1) Or (FindWindow(vbNullString, AU2100_MPTool_Title)) = 0 Then
            MPTester.Print "wait for MP Tool Ready"
            Call LoadMP_Click_AU2100
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
        Loop Until AlcorMPMessage = WM_MP_TEST_READY Or PassTime > 2 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                
        MPTester.Print "Ready Time="; PassTime
        If PassTime > 2 Then
            MPTester.TestResultLab = "Bin2:AU2100 MP Ready Fail"
            TestResult = "Bin2"
            Exit Sub
        End If
        
        OldTimer = Timer
        AlcorMPMessage = 0
        Call StartMP_Click_AU2100
        
        Do
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
                     
            PassTime = Timer - OldTimer
                   
        Loop Until AlcorMPMessage = WM_MP_FAIL _
                Or AlcorMPMessage = WM_MP_PASS _
                Or PassTime > 22 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                
        MPTester.Print "RW work Time="; PassTime
        
        '===========================================================
        '  MP Time Out Fail
        '===========================================================
                
        If PassTime > 22 Then
            winHwnd = FindWindow(vbNullString, AU2100_MPTool_Title)
            Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(1#)
                winHwnd = FindWindow(vbNullString, AU2100_MPTool_Title)
            Loop While winHwnd <> 0
            
            MPTester.TestResultLab = "Bin2:MP Time Out Fail"
            TestResult = "Bin2"
            Exit Sub
        End If
        
        If AlcorMPMessage <> WM_MP_PASS Then
            MPTester.TestResultLab = "Bin3:MP Fail"
            TestResult = "Bin3"
            Exit Sub
        End If
        
        Call MsecDelay(1#)
        
    End If  'End of bAU2100_OTP = True
    
    '===================================================================================================
    '===================================================================================================
    '===================================================================================================
    
    If (NewChipFlag = 1) Or (FindWindow(vbNullString, AU2100_AP_Title) = 0) Then
        MPTester.Print "wait for FT Tool Ready"
        NewChipFlag = 1
        Call LoadAP_Click_AU2100(OpenArg)
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
    Loop Until AlcorMPMessage = WM_FT_TEST_READY Or PassTime > 2 _
            Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                  
    Call Clear_Click_AU2100
    MPTester.Print "Ready Time="; PassTime
            
    If PassTime > 2 Then
        winHwnd = FindWindow(vbNullString, AU2100_AP_Title)
        If winHwnd <> 0 Then
            Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(1#)
                winHwnd = FindWindow(vbNullString, AU2100_AP_Title)
            Loop While winHwnd <> 0
        End If
    
        MPTester.TestResultLab = "Bin2:AP Ready Fail"
        TestResult = "Bin2"
        Exit Sub
    End If
    
    
    If Not (Inquiry_CK2_Connected) Then
        If Not (StartAP_Click_AU2100(WM_Connect_CK2, 0)) Then
            MPTester.TestResultLab = "Bin2: Connect CK2 Fail"
            TestResult = "Bin2"
            Exit Sub
        End If
    End If
    
    Call ReStart_Click_AU2100
    
'Flow 1:
    TempRes = 1
    
    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
    Call MsecDelay(0.1)
    'Device Output All GPIO Lo
    cardresult = DO_WritePort(card, Channel_P1A, &HEE)  'Open ENA、ENC(H)、Switch_En、Switch_Sel(H) 1110 1110
    'Call MsecDelay(0.01)
    If (StartAP_Click_AU2100(WM_SetOutput_0, 0)) Then
        Call MsecDelay(0.15)
        cardresutl = DI_ReadPort(card, Channel_P1B, GateVal)
        If GateVal <> &HF0 Then
            TempRes = 4
        End If
    Else
        MPTester.Print "Set Output-L Fail"
        TempRes = 4
    End If
    
'Flow 2:
    'Device Output All GPIO Hi
    If TempRes = 1 Then
        If StartAP_Click_AU2100(WM_SetOutput_1, 0) Then
            Call MsecDelay(0.45)
            cardresutl = DI_ReadPort(card, Channel_P1B, GateVal)
            If GateVal <> &HF Then
                TempRes = 4
            End If
        Else
            MPTester.Print "Set Output-H Fail"
            TempRes = 4
        End If
    End If
    
'Flow3:
    'Device change to Input0 mode
    If TempRes = 1 Then
        cardresult = DO_WritePort(card, Channel_P1A, &HE6)  'Open ENA、ENC(H)、Switch_En、Switch_Sel(L) 1110 0110
        'Call MsecDelay(0.01)
        If (StartAP_Click_AU2100(WM_SetInput_0, 0)) Then
            Call MsecDelay(0.01)
            
            If Not (StartAP_Click_AU2100(WM_SetRead, 0)) Then
                MPTester.Print "Read Device GPIO-Low Fail"
                TempRes = 5
            End If
            
        Else
            MPTester.Print "Set Input-0 Fail"
            TempRes = 5
        End If
    End If

'Flow4:
'Device change to Input0 mode
    If TempRes = 1 Then
        cardresult = DO_WritePort(card, Channel_P1A, &HE2)  'Open ENA、ENC(L)、Switch_En、Switch_Sel(L) 1110 0010
        'Call MsecDelay(0.01)
        If (StartAP_Click_AU2100(WM_SetInput_1, 0)) Then
            Call MsecDelay(0.01)
            
            If Not (StartAP_Click_AU2100(WM_SetRead, 1)) Then
                MPTester.Print "Read Device GPIO-H Fail"
                TempRes = 5
            End If
            
        Else
            MPTester.Print "Set Input-1 Fail"
            TempRes = 5
        End If
    End If
    
    
    If bAU2100_OTP Then
        If (Inquiry_CK2_Connected) Then
            Call StartAP_Click_AU2100(WM_DisConnect_CK2, 0)
        End If
    End If
    
    
AU2100TestEndLabel:

    cardresult = DO_WritePort(card, Channel_P1A, &HFF)
    Call PWROff_Click_AU2100
    
    If TempRes = 4 Then
        MPTester.TestResultLab = "Bin4: Output Fail"
        TestResult = "Bin4"
        AU2100_ContiFail = AU2100_ContiFail + 1
    ElseIf TempRes = 5 Then
        MPTester.TestResultLab = "Bin5: Input Fail"
        TestResult = "Bin5"
        AU2100_ContiFail = AU2100_ContiFail + 1
    ElseIf TempRes = 1 Then
        MPTester.TestResultLab = "PASS"
        TestResult = "PASS"
        AU2100_ContiFail = 0
    End If
    
    
                            
End Sub

Public Sub AU2101TestSub()
 
'This code copy from AU2100TestSub
'Purpose replace AU2101 AP path
 
Dim OldTimer
Dim PassTime
Dim rt2
Dim mMsg As MSG
Dim TempRes As Byte
Dim GateVal As Long


    'P1A(Output)    P1B(Input)
    '8765 4321      8765 4321
    
    '   S SEEE      GGGG GGGG
    '   W WNNN      AAAA AAAA
    '   I I|||      TTTT TTTT
    '   T TCBA      EEEE EEEE
    '   C C
    '   H H         ---- OOOO
    '               AAAA RRRR
    '   E S         NNNN
    '   N E         DDDD
    '   A L
    
    'Switch ENA : 0: active
    'Switch Sel : 0: Device Input; 1: Device Output
    

    MPTester.TestResultLab = ""
    
    If PCI7248InitFinish = 0 Then
        PCI7248Exist
    End If
    
    result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
    AlcorMPMessage = 0
    NewChipFlag = 0
    
    
    If OldChipName <> ChipName Then
        ' reset program
        winHwnd = FindWindow(vbNullString, AU2101_AP_Title)
        If winHwnd <> 0 Then
            Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(1#)
                winHwnd = FindWindow(vbNullString, AU2101_AP_Title)
            Loop While winHwnd <> 0
        End If
        
        winHwnd = FindWindow(vbNullString, AU2101_MPTool_Title)
        If winHwnd <> 0 Then
            Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(1#)
                winHwnd = FindWindow(vbNullString, AU2101_MPTool_Title)
            Loop While winHwnd <> 0
        End If
        
        winHwnd = FindWindow(vbNullString, AU2101_FPTool_Title)
        If winHwnd <> 0 Then
            Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(1#)
                winHwnd = FindWindow(vbNullString, AU2101_FPTool_Title)
            Loop While winHwnd <> 0
        End If
        
        NewChipFlag = 1
    End If
              
    OldChipName = ChipName
    
    '===================================================================================================
    '============================================= MP Flow =============================================
    '===================================================================================================
    
    If (bAU2100_OTP) Then       'MP + FT2
    
        cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        
        If (Inquiry_CK2_Connected_AU2101) Then
            Call StartAP_Click_AU2101(WM_DisConnect_CK2, 1)
        End If
        
        If (NewChipFlag = 1) Or (FindWindow(vbNullString, AU2101_MPTool_Title)) = 0 Then
            MPTester.Print "wait for MP Tool Ready"
            Call LoadMP_Click_AU2101
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
        Loop Until AlcorMPMessage = WM_MP_TEST_READY Or PassTime > 2 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                
        MPTester.Print "Ready Time="; PassTime
        If PassTime > 2 Then
            MPTester.TestResultLab = "Bin2:AU2101 MP Ready Fail"
            TestResult = "Bin2"
            Exit Sub
        End If
        
        'Connect CK2
        OldTimer = Timer
        AlcorMPMessage = 0
        Call MPConnectCK2_Click_AU2101
        
        Do
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
                     
            PassTime = Timer - OldTimer
                   
        Loop Until AlcorMPMessage = WM_MP_Connect_PASS _
                Or AlcorMPMessage = WM_MP_Connect_FAIL _
                Or PassTime > 6 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
        
        If (PassTime > 6) Or (AlcorMPMessage <> WM_MP_Connect_PASS) Then
            winHwnd = FindWindow(vbNullString, AU2101_MPTool_Title)
            Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(1#)
                winHwnd = FindWindow(vbNullString, AU2101_MPTool_Title)
            Loop While winHwnd <> 0
            
            MPTester.TestResultLab = "Bin2:Connect CK2 Time Out Fail"
            TestResult = "Bin2"
            Exit Sub
        End If
        MsecDelay (0.1)
        
        'Start MP
        OldTimer = Timer
        AlcorMPMessage = 0
        Call StartMP_Click_AU2101
        
        Do
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
                     
            PassTime = Timer - OldTimer
                   
        Loop Until AlcorMPMessage = WM_MP_FAIL _
                Or AlcorMPMessage = WM_MP_PASS _
                Or PassTime > 22 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                
        MPTester.Print "RW work Time="; PassTime
        
        '===========================================================
        '  MP Time Out Fail
        '===========================================================
                
        If PassTime > 22 Then
            winHwnd = FindWindow(vbNullString, AU2101_MPTool_Title)
            Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(1#)
                winHwnd = FindWindow(vbNullString, AU2101_MPTool_Title)
            Loop While winHwnd <> 0
            
            MPTester.TestResultLab = "Bin2:MP Time Out Fail"
            TestResult = "Bin2"
            Exit Sub
        End If
        
        If AlcorMPMessage <> WM_MP_PASS Then
            MPTester.TestResultLab = "Bin3:MP Fail"
            TestResult = "Bin3"
            Exit Sub
        End If
        MsecDelay (0.1)
        
        'Disconnect CK2
        OldTimer = Timer
        AlcorMPMessage = 0
        Call MPDisConnectCK2_Click_AU2101
        
        Do
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
                     
            PassTime = Timer - OldTimer
                   
        Loop Until AlcorMPMessage = WM_MP_DisConnect_PASS _
                Or AlcorMPMessage = WM_MP_DisConnect_FAIL _
                Or PassTime > 3 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
        
        If (PassTime > 3) Or (AlcorMPMessage <> WM_MP_DisConnect_PASS) Then
            winHwnd = FindWindow(vbNullString, AU2101_MPTool_Title)
            Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(1#)
                winHwnd = FindWindow(vbNullString, AU2101_MPTool_Title)
            Loop While winHwnd <> 0
            
            MPTester.TestResultLab = "Bin2:DisConnect CK2 Time Out Fail"
            TestResult = "Bin2"
            Exit Sub
        End If
        
        Call MsecDelay(0.2)
        
    End If  'End of bAU2100_OTP = True
    
    '===================================================================================================
    '===================================================================================================
    '===================================================================================================
    
    If (NewChipFlag = 1) Or (FindWindow(vbNullString, AU2101_AP_Title) = 0) Then
        MPTester.Print "wait for FT Tool Ready"
        NewChipFlag = 1
        Call LoadAP_Click_AU2101(OpenArg)
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
    Loop Until AlcorMPMessage = WM_FT_TEST_READY Or PassTime > 2 _
            Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                  
    Call Clear_Click_AU2101
    MPTester.Print "Ready Time="; PassTime
            
    If PassTime > 2 Then
        winHwnd = FindWindow(vbNullString, AU2101_AP_Title)
        If winHwnd <> 0 Then
            Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(1#)
                winHwnd = FindWindow(vbNullString, AU2101_AP_Title)
            Loop While winHwnd <> 0
        End If
    
        MPTester.TestResultLab = "Bin2:AP Ready Fail"
        TestResult = "Bin2"
        Exit Sub
    End If
    
    
    If Not (Inquiry_CK2_Connected_AU2101) Then
        If Not (StartAP_Click_AU2101(WM_Connect_CK2, 0)) Then
            MPTester.TestResultLab = "Bin2: Connect CK2 Fail"
            TestResult = "Bin2"
            Exit Sub
        End If
    End If
    
    
    Call ReStart_Click_AU2101
    Call MsecDelay(0.2)
    Call SetPinNumber_AU2101
    
    
'Flow 1:
    TempRes = 1
    
    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
    Call MsecDelay(0.1)
    'Device Output All GPIO Lo
    cardresult = DO_WritePort(card, Channel_P1A, &HEE)  'Open ENA、ENC(H)、Switch_En、Switch_Sel(H) 1110 1110
    'Call MsecDelay(0.01)
    If (StartAP_Click_AU2101(WM_SetOutput_0, 0)) Then
        Call MsecDelay(0.15)
        cardresutl = DI_ReadPort(card, Channel_P1B, GateVal)
        If GateVal <> &HF0 Then
            TempRes = 4
        End If
    Else
        MPTester.Print "Set Output-L Fail"
        TempRes = 4
    End If
    
'Flow 2:
    'Device Output All GPIO Hi
    If TempRes = 1 Then
        If StartAP_Click_AU2101(WM_SetOutput_1, 0) Then
            Call MsecDelay(0.01)
            cardresutl = DI_ReadPort(card, Channel_P1B, GateVal)
            If GateVal <> &HF Then
                TempRes = 4
            End If
        Else
            MPTester.Print "Set Output-H Fail"
            TempRes = 4
        End If
    End If
    
'Flow3:
    'Device change to Input0 mode
    If TempRes = 1 Then
        cardresult = DO_WritePort(card, Channel_P1A, &HE6)  'Open ENA、ENC(H)、Switch_En、Switch_Sel(L) 1110 0110
        'Call MsecDelay(0.01)
        If (StartAP_Click_AU2101(WM_SetInput_0, 0)) Then
            Call MsecDelay(0.01)
            
            If Not (StartAP_Click_AU2101(WM_SetRead, 0)) Then
                MPTester.Print "Read Device GPIO-Low Fail"
                TempRes = 5
            End If
            
        Else
            MPTester.Print "Set Input-0 Fail"
            TempRes = 5
        End If
    End If

'Flow4:
'Device change to Input1 mode
    If TempRes = 1 Then
        Call SetPinNumber_AU2101
        cardresult = DO_WritePort(card, Channel_P1A, &HE2)  'Open ENA、ENC(L)、Switch_En、Switch_Sel(L) 1110 0010
        'Call MsecDelay(0.01)
        If (StartAP_Click_AU2101(WM_SetInput_1, 1)) Then
            Call MsecDelay(0.01)
            
            If Not (StartAP_Click_AU2101(WM_SetRead, 1)) Then
                MPTester.Print "Read Device GPIO-H Fail"
                TempRes = 5
            End If
            
        Else
            MPTester.Print "Set Input-1 Fail"
            TempRes = 5
        End If
    End If
    
    
    If bAU2100_OTP Then
        If (Inquiry_CK2_Connected_AU2101) Then
            Call StartAP_Click_AU2101(WM_DisConnect_CK2, 0)
        End If
    End If
    
    
AU2101TestEndLabel:

    cardresult = DO_WritePort(card, Channel_P1A, &HFF)
    Call PWROff_Click_AU2101
    
    If TempRes = 4 Then
        MPTester.TestResultLab = "Bin4: Output Fail"
        TestResult = "Bin4"
        AU2100_ContiFail = AU2100_ContiFail + 1
    ElseIf TempRes = 5 Then
        MPTester.TestResultLab = "Bin5: Input Fail"
        TestResult = "Bin5"
        AU2100_ContiFail = AU2100_ContiFail + 1
    ElseIf TempRes = 1 Then
        MPTester.TestResultLab = "PASS"
        TestResult = "PASS"
        AU2100_ContiFail = 0
    End If
    
    
                            
End Sub

Public Sub AU2101_FPXXXX_Sub()
 
Dim OldTimer
Dim PassTime
Dim rt2
Dim mMsg As MSG
Dim TempRes As Byte
Dim GateVal As Long


    'P1A(Output)    P1B(Input)
    '8765 4321      8765 4321
    
    '   S SEEE      GGGG GGGG
    '   W WNNN      AAAA AAAA
    '   I I|||      TTTT TTTT
    '   T TCBA      EEEE EEEE
    '   C C
    '   H H         ---- OOOO
    '               AAAA RRRR
    '   E S         NNNN
    '   N E         DDDD
    '   A L
    
    'Switch ENA : 0: active
    'Switch Sel : 0: Device Input; 1: Device Output
    

    MPTester.TestResultLab = ""
    
    If PCI7248InitFinish = 0 Then
        PCI7248Exist
    End If
    
'    result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
'    AlcorMPMessage = 0
'    NewChipFlag = 0
'
'
    If OldChipName <> ChipName Then
        ' reset program
        winHwnd = FindWindow(vbNullString, AU2101_AP_Title)
        If winHwnd <> 0 Then
            Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(1#)
                winHwnd = FindWindow(vbNullString, AU2101_AP_Title)
            Loop While winHwnd <> 0
        End If

        winHwnd = FindWindow(vbNullString, AU2101_MPTool_Title)
        If winHwnd <> 0 Then
            Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(1#)
                winHwnd = FindWindow(vbNullString, AU2101_MPTool_Title)
            Loop While winHwnd <> 0
        End If

        winHwnd = FindWindow(vbNullString, AU2101_FPTool_Title)
        If winHwnd <> 0 Then
            Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(1#)
                winHwnd = FindWindow(vbNullString, AU2101_FPTool_Title)
            Loop While winHwnd <> 0
        End If

        NewChipFlag = 1
    End If

    OldChipName = ChipName
    TempRes = 1
'
'
'    If (NewChipFlag = 1) Or (FindWindow(vbNullString, AU2101_AP_Title) = 0) Then
'        MPTester.Print "wait for FT Tool Ready"
'        NewChipFlag = 1
'        Call LoadAP_Click_AU2101(OpenArg)
'    End If
'
'    OldTimer = Timer
'    AlcorMPMessage = 0
'
'    Do
'        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
'            AlcorMPMessage = mMsg.message
'            TranslateMessage mMsg
'            DispatchMessage mMsg
'        End If
'        PassTime = Timer - OldTimer
'    Loop Until AlcorMPMessage = WM_FT_TEST_READY Or PassTime > 2 _
'            Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
'
'    Call Clear_Click_AU2101
'
'    MPTester.Print "Ready Time="; PassTime
'
'    If PassTime > 2 Then
'        winHwnd = FindWindow(vbNullString, AU2101_AP_Title)
'        If winHwnd <> 0 Then
'            Do
'                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
'                Call MsecDelay(1#)
'                winHwnd = FindWindow(vbNullString, AU2101_AP_Title)
'            Loop While winHwnd <> 0
'        End If
'
'        MPTester.TestResultLab = "Bin2:AP Ready Fail"
'        TestResult = "Bin2"
'        Exit Sub
'    End If
'
'
'    If Not (Inquiry_CK2_Connected_AU2101) Then
'        If Not (StartAP_Click_AU2101(WM_Connect_CK2, 0)) Then
'            MPTester.TestResultLab = "Bin2: Connect CK2 Fail"
'            TestResult = "Bin2"
'            Exit Sub
'        End If
'    End If
'
'
'    Call ReStart_Click_AU2101
'    Call MsecDelay(0.2)
'    Call SetPinNumber_AU2101
'
'
''Flow 1:
'    TempRes = 1
'
'    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
'    Call MsecDelay(0.1)
'    'Device Output All GPIO Lo
'    cardresult = DO_WritePort(card, Channel_P1A, &HEE)  'Open ENA、ENC(H)、Switch_En、Switch_Sel(H) 1110 1110
'    'Call MsecDelay(0.01)
'    If (StartAP_Click_AU2101(WM_SetOutput_0, 0)) Then
'        Call MsecDelay(0.15)
'        cardresutl = DI_ReadPort(card, Channel_P1B, GateVal)
'        If GateVal <> &HF0 Then
'            TempRes = 4
'        End If
'    Else
'        MPTester.Print "Set Output-L Fail"
'        TempRes = 4
'    End If
'
''Flow 2:
'    'Device Output All GPIO Hi
'    If TempRes = 1 Then
'        If StartAP_Click_AU2101(WM_SetOutput_1, 0) Then
'            Call MsecDelay(0.01)
'            cardresutl = DI_ReadPort(card, Channel_P1B, GateVal)
'            If GateVal <> &HF Then
'                TempRes = 4
'            End If
'        Else
'            MPTester.Print "Set Output-H Fail"
'            TempRes = 4
'        End If
'    End If
'
''Flow3:
'    'Device change to Input0 mode
'    If TempRes = 1 Then
'        cardresult = DO_WritePort(card, Channel_P1A, &HE6)  'Open ENA、ENC(H)、Switch_En、Switch_Sel(L) 1110 0110
'        'Call MsecDelay(0.01)
'        If (StartAP_Click_AU2101(WM_SetInput_0, 0)) Then
'            Call MsecDelay(0.01)
'
'            If Not (StartAP_Click_AU2101(WM_SetRead, 0)) Then
'                MPTester.Print "Read Device GPIO-Low Fail"
'                TempRes = 5
'            End If
'
'        Else
'            MPTester.Print "Set Input-0 Fail"
'            TempRes = 5
'        End If
'    End If
'
''Flow4:
''Device change to Input1 mode
'    If TempRes = 1 Then
'        Call SetPinNumber_AU2101
'        cardresult = DO_WritePort(card, Channel_P1A, &HE2)  'Open ENA、ENC(L)、Switch_En、Switch_Sel(L) 1110 0010
'        'Call MsecDelay(0.01)
'        If (StartAP_Click_AU2101(WM_SetInput_1, 1)) Then
'            Call MsecDelay(0.01)
'
'            If Not (StartAP_Click_AU2101(WM_SetRead, 1)) Then
'                MPTester.Print "Read Device GPIO-H Fail"
'                TempRes = 5
'            End If
'
'        Else
'            MPTester.Print "Set Input-1 Fail"
'            TempRes = 5
'        End If
'    End If
'
'
'    If bAU2100_OTP Then
'        If (Inquiry_CK2_Connected_AU2101) Then
'            Call StartAP_Click_AU2101(WM_DisConnect_CK2, 0)
'        End If
'    End If
'
'    If TempRes <> 1 Then
'        GoTo AU2101TestEndLabel
'    End If
    
    
    '===================================================================================================
    '================================ MP for Customer's FW Flow ========================================
    '===================================================================================================
    
    If (bAU2100_OTP) Then
    
        cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
        
'        If (Inquiry_CK2_Connected_AU2101) Then
'            Call StartAP_Click_AU2101(WM_DisConnect_CK2, 1)
'        End If
        
        If (NewChipFlag = 1) Or (FindWindow(vbNullString, AU2101_FPTool_Title)) = 0 Then
            MPTester.Print "wait for MP Tool Ready"
            Call LoadMP_Click_AU2101_FPXXXX(AU2101_DllName, AU2101_BinName)
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
        Loop Until AlcorMPMessage = WM_MP_TEST_READY Or PassTime > 2 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                
        MPTester.Print "Ready Time="; PassTime
        If PassTime > 2 Then
            MPTester.TestResultLab = "Bin2:AU2101 MP Ready Fail"
            TestResult = "Bin2"
            Exit Sub
        End If
        
        'Connect CK2
        OldTimer = Timer
        AlcorMPMessage = 0
        Call MPConnectCK2_Click_AU2101_FPXXXX
        
        Do
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
                     
            PassTime = Timer - OldTimer
                   
        Loop Until AlcorMPMessage = WM_MP_Connect_PASS _
                Or AlcorMPMessage = WM_MP_Connect_FAIL _
                Or PassTime > 6 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
        
        If (PassTime > 6) Or (AlcorMPMessage <> WM_MP_Connect_PASS) Then
            winHwnd = FindWindow(vbNullString, AU2101_FPTool_Title)
            Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(1#)
                winHwnd = FindWindow(vbNullString, AU2101_FPTool_Title)
            Loop While winHwnd <> 0
            
            MPTester.TestResultLab = "Bin2:Connect CK2 Time Out Fail"
            TestResult = "Bin2"
            Exit Sub
        End If
        MsecDelay (0.1)
        
        'Start MP
        OldTimer = Timer
        AlcorMPMessage = 0
        Call StartMP_Click_AU2101_FPXXXX
        
        Do
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
                     
            PassTime = Timer - OldTimer
                   
        Loop Until AlcorMPMessage = WM_MP_FAIL _
                Or AlcorMPMessage = WM_MP_PASS _
                Or PassTime > 22 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                
        MPTester.Print "RW work Time="; PassTime
        
        '===========================================================
        '  MP Time Out Fail
        '===========================================================
                
        If PassTime > 22 Then
            winHwnd = FindWindow(vbNullString, AU2101_FPTool_Title)
            Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(1#)
                winHwnd = FindWindow(vbNullString, AU2101_FPTool_Title)
            Loop While winHwnd <> 0
            
            MPTester.TestResultLab = "Bin2:MP Time Out Fail"
            TestResult = "Bin2"
            Exit Sub
        End If
        
        If AlcorMPMessage <> WM_MP_PASS Then
            MPTester.TestResultLab = "Bin3:MP Fail"
            TestResult = "Bin3"
            Exit Sub
        End If
        MsecDelay (0.1)
        
        '====================================================================
        '=============== Check FW Version From Device =======================
        '====================================================================
        Call MPReadFWVersion_Click_AU2101_FPXXXX
        
        MsecDelay (1.5)
        
        

        If (MPReadFWVersion_Click_AU2101(WM_MP_ReadFW_CusID) <> Target_DevID) Or _
           (MPReadFWVersion_Click_AU2101(WM_MP_ReadFW_DevID) <> Target_DevID) Or _
           (MPReadFWVersion_Click_AU2101(WM_MP_ReadFW_ResNum) <> Target_ResNum) Or _
           (MPReadFWVersion_Click_AU2101(WM_MP_ReadFW_ExtFW) <> Target_ExtFW) Or _
           (MPReadFWVersion_Click_AU2101(WM_MP_ReadFW_SrcFW) <> Target_SrcFW) Or _
           (MPReadFWVersion_Click_AU2101(WM_MP_ReadFW_FWLen) <> Target_FWLen) Or _
           (MPReadFWVersion_Click_AU2101(WM_MP_ReadFW_CrcVal) <> Target_CrcVal) Then
           
            MPTester.TestResultLab = "Bin3:MP Fail"
            TestResult = "Bin3"
            
            Exit Sub
        End If

'        If (MPReadFWVersion_Click_AU2101(WM_MP_ReadFW_CrcVal) <> Target_CrcVal) Then
'            a = a
'        End If
        
        'Disconnect CK2
        OldTimer = Timer
        AlcorMPMessage = 0
        Call MPDisConnectCK2_Click_AU2101_FPXXXX
        
        Do
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
                     
            PassTime = Timer - OldTimer
                   
        Loop Until AlcorMPMessage = WM_MP_DisConnect_PASS _
                Or AlcorMPMessage = WM_MP_DisConnect_FAIL _
                Or PassTime > 3 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
        
        If (PassTime > 3) Or (AlcorMPMessage <> WM_MP_DisConnect_PASS) Then
            winHwnd = FindWindow(vbNullString, AU2101_FPTool_Title)
            Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(1#)
                winHwnd = FindWindow(vbNullString, AU2101_FPTool_Title)
            Loop While winHwnd <> 0
            
            MPTester.TestResultLab = "Bin2:DisConnect CK2 Time Out Fail"
            TestResult = "Bin2"
            Exit Sub
        End If
        
        Call MsecDelay(0.2)
        
    End If  'End of bAU2100_OTP = True
    
    '===================================================================================================
    '===================================================================================================
    '===================================================================================================
    
    
AU2101TestEndLabel:

    cardresult = DO_WritePort(card, Channel_P1A, &HFF)
    Call PWROff_Click_AU2101
    
'    If TempRes = 4 Then
'        MPTester.TestResultLab = "Bin4: Output Fail"
'        TestResult = "Bin4"
'        AU2100_ContiFail = AU2100_ContiFail + 1
'    ElseIf TempRes = 5 Then
'        MPTester.TestResultLab = "Bin5: Input Fail"
'        TestResult = "Bin5"
'        AU2100_ContiFail = AU2100_ContiFail + 1
'    ElseIf TempRes = 1 Then
    If TempRes = 1 Then
        MPTester.TestResultLab = "PASS"
        TestResult = "PASS"
        AU2100_ContiFail = 0
    End If
    
                            
End Sub

Public Sub AU2100ProgSub()
 
Dim OldTimer
Dim PassTime
Dim rt2
Dim mMsg As MSG
Dim TempRes As Byte
Dim GateVal As Long

'20130923 Update for new version MP Tool(integrate ini,bin file)

    'P1A(Output)    P1B(Input)
    '8765 4321      8765 4321
    
    '   S SEEE      GGGG GGGG
    '   W WNNN      AAAA AAAA
    '   I I|||      TTTT TTTT
    '   T TCBA      EEEE EEEE
    '   C C
    '   H H         ---- OOOO
    '               AAAA RRRR
    '   E S         NNNN
    '   N E         DDDD
    '   A L
    
    'Switch ENA : 0: active
    'Switch Sel : 0: Device Input; 1: Device Output
    

    MPTester.TestResultLab = ""
    
    If PCI7248InitFinish = 0 Then
        PCI7248Exist
    End If
    
    result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
    Call MsecDelay(0.02)
    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
    AlcorMPMessage = 0
    NewChipFlag = 0
    
    
    If OldChipName <> ChipName Then
        ' reset program
        winHwnd = FindWindow(vbNullString, AU2100_AP_Title)
        If winHwnd <> 0 Then
            Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(1#)
                winHwnd = FindWindow(vbNullString, AU2100_AP_Title)
            Loop While winHwnd <> 0
        End If
        
        winHwnd = FindWindow(vbNullString, AU2100_MPTool_Title)
        If winHwnd <> 0 Then
            Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(1#)
                winHwnd = FindWindow(vbNullString, AU2100_MPTool_Title)
            Loop While winHwnd <> 0
        End If
        
        winHwnd = FindWindow(vbNullString, AU2100_NewMPTool_Title)
        If winHwnd <> 0 Then
            Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(1#)
                winHwnd = FindWindow(vbNullString, AU2100_NewMPTool_Title)
            Loop While winHwnd <> 0
        End If
        
        NewChipFlag = 1
    End If
              
    OldChipName = ChipName
    
    '===================================================================================================
    '============================================= MP Flow =============================================
    '===================================================================================================
    
    
    cardresult = DO_WritePort(card, Channel_P1A, &HFF)
    
    If (NewChipFlag = 1) Or (FindWindow(vbNullString, AU2100_NewMPTool_Title)) = 0 Then
        MPTester.Print "wait for MP Tool Ready"
        Call LoadNewMP_Click_AU2100
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
    Loop Until AlcorMPMessage = WM_MP_TEST_READY Or PassTime > 2 _
            Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
            
    MPTester.Print "Ready Time="; PassTime
    If PassTime > 2 Then
        MPTester.TestResultLab = "Bin2:AU2100 MP Ready Fail"
        TestResult = "Bin2"
        Exit Sub
    End If
    
    OldTimer = Timer
    AlcorMPMessage = 0
    Call StartNewMP_Click_AU2100
    
    Do
        If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
            AlcorMPMessage = mMsg.message
            TranslateMessage mMsg
            DispatchMessage mMsg
        End If
                 
        PassTime = Timer - OldTimer
               
    Loop Until AlcorMPMessage = WM_MP_FAIL _
            Or AlcorMPMessage = WM_MP_PASS _
            Or PassTime > 15 _
            Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
            
    MPTester.Print "RW work Time="; PassTime
    
    '===========================================================
    '  MP Time Out Fail
    '===========================================================
            
    If PassTime > 12 Then
        winHwnd = FindWindow(vbNullString, AU2100_NewMPTool_Title)
        Do
            rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
            Call MsecDelay(1#)
            winHwnd = FindWindow(vbNullString, AU2100_NewMPTool_Title)
        Loop While winHwnd <> 0
        
        MPTester.TestResultLab = "Bin2:MP Time Out Fail"
        TestResult = "Bin2"
        Exit Sub
    End If
    
    If AlcorMPMessage <> WM_MP_PASS Then
        MPTester.TestResultLab = "Bin3:MP Fail"
        TestResult = "Bin3"
        Exit Sub
    End If
        
    
    
AU2100TestEndLabel:

    cardresult = DO_WritePort(card, Channel_P1A, &HFF)
    
    If AlcorMPMessage = WM_MP_PASS Then
        MPTester.TestResultLab = "PASS"
        TestResult = "PASS"
    End If
    
                            
End Sub
