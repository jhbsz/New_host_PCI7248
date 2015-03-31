Attribute VB_Name = "AU9520Test"
Option Explicit

Public FindSmartCardResult As Byte
Public SmartCardTestResult As Byte

Public Type msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Public Declare Function WaitMessage Lib "user32" () As Boolean
Public Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Public Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As msg) As Long
Public Declare Function TranslateMessage Lib "user32" (lpMsg As msg) As Long

Public Const SW_SHOW = 5

Public Const PM_NOREMOVE = &H0
Public Const PM_REMOVE = &H1

Const SWP_NOMOVE = &H2 '不更動目前視窗位置
Const SWP_NOSIZE = &H1 '不更動目前視窗大小
Const HWND_TOPMOST = -1 '設定為最上層
Const HWND_NOTOPMOST = -2 '取消最上層設定
Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4

Public Const WM_USER = &H400
Public Const WM_COM_START_TEST = WM_USER + &H700
Public Const WM_COM_TEST_PASS = WM_USER + &H710
Public Const WM_COM_TEST_FAIL = WM_USER + &H720
Public Const WM_COM_TEST_UNKNOWN = WM_USER + &H730
Public Const WM_COM_TEST_CLOSE = WM_USER + &H740
Public Const WM_NO_ECHO_WPAR_FLAG = WM_USER + &H99

Public Sub AU9520TestGPIB()


  Tester.Cls
                  If ChipName = "AU9520V4" Then
                Tester.Print "====== 4.2  V test"
                  Call PowerSet(42)
                  Else
               Tester.Print "====== 5.0 V test"
                  Call PowerSet(50)
                  
                  End If
                  
                 
                  rv0 = AU9520_2Slot
                 
                 Call LabelMenu(0, rv0, 1)
                 
                 If ChipName = "AU9520V5" Or ChipName = "AU9520V4" Or ChipName = "AU9520ALF20" Then
                        If rv0 = 1 Then
                              If ChipName = "AU9520V4" Then
                                Call PowerSet(42)
                                Else
                                Call PowerSet(50)
                               End If
                  
                            rv2 = AU9520JJMode
                         
                        End If
                        Call LabelMenu(2, rv2, rv0)
                  Else
                       rv2 = 1
                        
                  End If
                 
              
                
               
                 
                 
                 If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv1 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv1 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
                    TestResult = "CF_RF"
                ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                 ElseIf rv3 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv3 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv0 * rv2 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub

Public Sub AU9520FLF21TestSub()


  Tester.Cls
                   
                  
                    rv0 = AU9520FLF21SingleSlot
                 
                 
                 Call LabelMenu(0, rv0, 1)
                 
                 
                If rv0 = 1 Then
                              
                             rv2 = AU9520_2Slot
                         
                End If
                 Call LabelMenu(2, rv2, rv0)
                  
                 
              
                
               
                 
                 
                 If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv1 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv1 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
                    TestResult = "CF_RF"
                ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                 ElseIf rv3 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv3 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv0 * rv2 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub
Public Sub AU9520FLF20TestSub()


  Tester.Cls
                   
                  
                 
                   rv0 = AU9520_2Slot
                 
                 Call LabelMenu(0, rv0, 1)
                 
                 
                If rv0 = 1 Then
                              
                            rv2 = AU9520FLSingleSlot
                         
                End If
                 Call LabelMenu(2, rv2, rv0)
                  
                 
              
                
               
                 
                 
                 If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv1 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv1 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
                    TestResult = "CF_RF"
                ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                 ElseIf rv3 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv3 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv0 * rv2 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub

Public Sub AU9520GLF20TestSub()


  Tester.Cls
                   
                  
                 
                   rv0 = AU9520_2Slot
                 
                 Call LabelMenu(0, rv0, 1)
                 
                 
                If rv0 = 1 Then
                              
                            rv2 = AU9520GLSingleSlot
                         
                End If
                 Call LabelMenu(2, rv2, rv0)
                  
                 
              
                
               
                 
                 
                 If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv1 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv1 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
                    TestResult = "CF_RF"
                ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                 ElseIf rv3 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv3 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv0 * rv2 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub

Public Sub AU9520TestNOGPIB()


  Tester.Cls
                  If ChipName = "AU9520V4" Then
                 ' Print "====== 4.2  V test"
                  Call PowerSet(42)
                  Else
                  Tester.Print "====== NO GPIB Test"
                  'Call PowerSet(50)
                  
                  End If
                  
                 
                  rv0 = AU9520_2Slot
                 
                 Call LabelMenu(0, rv0, 1)
                 
                 If ChipName = "AU9520V51" Or ChipName = "AU9520V4" Or ChipName = "AU9520ALF21" Then
                        If rv0 = 1 Then
                              If ChipName = "AU9520V4" Then
                                Call PowerSet(42)
                                Else
                             '   Call PowerSet(50)
                               End If
                  
                            rv2 = AU9520JJMode
                         
                        End If
                        Call LabelMenu(2, rv2, rv0)
                  Else
                       rv2 = 1
                        
                  End If
                 
              
                
               
                 
                 
                 If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv1 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv1 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
                    TestResult = "CF_RF"
                ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                 ElseIf rv3 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv3 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv0 * rv2 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub


Public Sub AU9525ALF20TestNOGPIB()


  Tester.Cls
                  If ChipName = "AU9520V4" Then
                 ' Print "====== 4.2  V test"
                  Call PowerSet(42)
                  Else
                  Tester.Print "====== NO GPIB Test"
                  'Call PowerSet(50)
                  
                  End If
                  
                 
                  rv0 = AU9520_2Slot
                 
                 Call LabelMenu(0, rv0, 1)
                 
                 
              
                
               
                 
                 
                 If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv1 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv1 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
                    TestResult = "CF_RF"
                ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                 ElseIf rv3 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv3 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv0 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub
Public Sub AU9525CLF20TestSub()
Dim tmpResult As Byte
Dim tmpStr As String
Dim TmpCount As Integer

    Tester.Cls
    tmpStr = ""
    rv0 = 0
    tmpResult = 0
    TmpCount = 0
        
    If PCI7248InitFinish = 0 Then
        PCI7248Exist
    End If
    
    CardResult = DO_WritePort(card, Channel_P1A, &H80)
    Call MsecDelay(1.5)
    
    Do
    
        tmpStr = GetDeviceName("vid_058f")
        Call MsecDelay(0.1)
        TmpCount = TmpCount + 1
        
    Loop Until (tmpStr <> "") Or (TmpCount > 10)
                
    If tmpStr = "" Then
        rv0 = 0
        Call LabelMenu(0, rv0, 1)
        GoTo AU9525CLF20TestResult
    Else
        'Tester.Print Mid(tmpStr, InStr(tmpStr, "vid_"), 8) & " , " & Mid(tmpStr, InStr(tmpStr, "pid_"), 8)
        rv0 = 1
        Call LabelMenu(0, rv0, 1)
    End If
    
                
    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
    Call MsecDelay(0.01)
    LightOn = CAndValue(LightOn, &H7)
    
    If (LightOn <> 0) Then
        rv0 = 2
        Tester.Print "GPO Fail! " & LightOn
        GoTo AU9525CLF20TestResult
    Else
        Tester.Print "GPO PASS! LightON: " & LightOn
    End If
    
    strReaderName = VenderString & " " & ProductString & " 0"   'slot 0
    udtReaderStates(0).szReader = strReaderName & vbNullChar
    Tester.txtmsg.Text = udtReaderStates(0).szReader
                        
    Tester.Print "Start to test slot0 !"
    
    Call StartTest
                         
    If SmartCardTestResult = 1 Then
                        
        SmartCardTestResult = 0  'Reset Test Result
        Tester.Print "slot0 test ok!"
                            
        strReaderName = VenderString & " " & ProductString & " 1"   'slot 1
        udtReaderStates(0).szReader = strReaderName & vbNullChar
        Tester.txtmsg.Text = udtReaderStates(0).szReader
                            
        Tester.Print vbCrLf & "Start to test slot1 !"
                           
        Call StartTest
        If SmartCardTestResult = 1 Then
            Tester.Print "slot1 test ok!"
            tmpResult = 1
        Else
            Tester.Print "slot1 test Fail!"
        End If
    Else
        
        Tester.Print "slot0 test Fail!"
    
    End If
                
    rv1 = tmpResult
                 
    Call LabelMenu(0, rv1, rv0)
                 
                 
AU9525CLF20TestResult:
                
                'CardResult = DO_WritePort(card, Channel_P1A, &H1)  'close power
                'Call MsecDelay(0.2)
    
                 If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv1 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv1 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
                    TestResult = "CF_RF"
                ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                 ElseIf rv3 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv3 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv0 * rv1 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub
Public Function AU9520JJMode() As Integer
On Error Resume Next
Dim tmpStr As String
tmpStr = ""

'(1) this is for JJ mode  --- one slot mode
     
FindSmartCardResult = 0
SmartCardTestResult = 0
                If PCI7248InitFinish = 0 Then
                      PCI7248Exist
                End If
                
                '3 bit: rom selection
                '2 bit: pwr enable
                '
                
                
               
                
                
                '=========================================
                '    POWER on
                '=========================================
                 ' result = DIO_PortConfig(card, Channel_P1CH, OUTPUT_PORT)
                 ' CardResult = DO_WritePort(card, Channel_P1CH, &H0)
                 '  CardResult = DO_WritePort(card, Channel_P1CL, &H4)  'Power Enable
                 
                    result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    result = DIO_PortConfig(card, Channel_P1A, OUTPUT_PORT)
                    CardResult = DO_WritePort(card, Channel_P1A, &H2)
                  
                  Call MsecDelay(0.3)
                  'CardResult = DO_WritePort(card, Channel_P1CL, &HF)
                  CardResult = DO_WritePort(card, Channel_P1A, &H7)
               '   CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.2)
              '   CardResult = DO_WritePort(card, Channel_P1CL, &HB)  'Power Enable
                 CardResult = DO_WritePort(card, Channel_P1A, &H5)  'Power Enable
                 Call MsecDelay(1.5)    'power on time
                 
                 If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
                 
            '  TmpStr = GetDeviceName("vid")
                 
             '     If TmpStr = "" Then
             '          AU9520JJMode = 0
             '          Exit Function
             '    End If
                 
                 
                 ' result = DIO_PortConfig(card, Channel_P1CH, INPUT_PORT)
                 result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                 '===========================================
                 'NO card test
                 '============================================
                '   CardResult = DO_ReadPort(card, Channel_P1CH, LightOFF)
                  Call MsecDelay(1)
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                   If CardResult <> 0 Then
                    MsgBox "Read card detect light off fail"
                    End
                   End If
                   
                  
                   
                 
                '===========================================
                 'NO card test
                 '============================================
               '   CardResult = DO_WritePort(card, Channel_P1CL, &HA)  'Power Enable + Slot0 enable
                    CardResult = DO_WritePort(card, Channel_P1A, &H4)
                   If CardResult <> 0 Then
                    MsgBox "Set card detect light on fail"
                    End
                   End If
                   Call MsecDelay(0.6)
                  '  CardResult = DO_ReadPort(card, Channel_P1CH, LightON)
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                   If CardResult <> 0 Then
                    MsgBox "Read light on fail"
                    End
                   End If
                   
                   
                  
                '===========================================
                 'R/W test
                 '============================================
                  
               ' If LightOFF <> 14 Or LightON <> 12 Then
               
               If Left(ChipName, 10) <> "AU9520ALF2" Then
                
               
                  If LightOff <> 254 Or LightOn <> 252 Then
                 Tester.Cls
                 Tester.Print "light fail"
                 Tester.Print "LightOFF="; LightOff
                  Tester.Print "LightON="; LightOn
                 
                   AU9520JJMode = 2
                   
                   
                   
                Else
                       
                         
                        Tester.Cls
                        
                        strReaderName = VenderString_AU9520JJ & " " & ProductString_AU9520JJ & " 0"   'slot 0
                        udtReaderStates(0).szReader = strReaderName & vbNullChar
                        Tester.txtmsg.Text = udtReaderStates(0).szReader
                        Tester.Print "Start to test slot0!"
                        Call StartTest
                        AU9520JJMode = SmartCardTestResult
                 End If
                 
             End If
             
   '=========for   ChipName = "AU9520ALF20"==================================================
                   
               If Left(ChipName, 10) = "AU9520ALF2" Then
                
               
                   If LightOff <> 255 Or LightOn <> 253 Then
               
                 Tester.Cls
                 Tester.Print "light fail"
                 Tester.Print "LightOFF="; LightOff
                  Tester.Print "LightON="; LightOn
                 
                   AU9520JJMode = 2
                   
                   
                   
                Else
                       
                         
                        Tester.Cls
                        
                        strReaderName = VenderString_AU9520JJ & " " & ProductString_AU9520JJ & " 0"   'slot 0
                        udtReaderStates(0).szReader = strReaderName & vbNullChar
                        Tester.txtmsg.Text = udtReaderStates(0).szReader
                        Tester.Print "Start to test slot0!"
                        Call StartTest
                        AU9520JJMode = SmartCardTestResult
                 End If
               End If
 End Function
 
 Public Function AU9520FLSingleSlot() As Integer
On Error Resume Next
Dim tmpStr As String
tmpStr = ""

'(1) this is for JJ mode  --- one slot mode
     
FindSmartCardResult = 0
SmartCardTestResult = 0
                If PCI7248InitFinish = 0 Then
                      PCI7248Exist
                End If
                
                '3 bit: rom selection
                '2 bit: pwr enable
                '
                
                
               
                
                
                '=========================================
                '    POWER on
                '=========================================
                 ' result = DIO_PortConfig(card, Channel_P1CH, OUTPUT_PORT)
                 ' CardResult = DO_WritePort(card, Channel_P1CH, &H0)
                 '  CardResult = DO_WritePort(card, Channel_P1CL, &H4)  'Power Enable
                 
                    result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    result = DIO_PortConfig(card, Channel_P1A, OUTPUT_PORT)
                    CardResult = DO_WritePort(card, Channel_P1A, &H2)
                  
                  Call MsecDelay(0.3)
                  'CardResult = DO_WritePort(card, Channel_P1CL, &HF)
                  CardResult = DO_WritePort(card, Channel_P1A, &H7)
               '   CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.2)
              '   CardResult = DO_WritePort(card, Channel_P1CL, &HB)  'Power Enable
                 CardResult = DO_WritePort(card, Channel_P1A, &H81)  'Power Enable
                 Call MsecDelay(1.5)    'power on time
                 
                 If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
                 
               tmpStr = GetDeviceName("vid_1b0e")
                 
                   If tmpStr = "" Then
                        AU9520FLSingleSlot = 0
                        Exit Function
                 End If
                 
                 
                 ' result = DIO_PortConfig(card, Channel_P1CH, INPUT_PORT)
                 result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                 '===========================================
                 'NO card test
                 '============================================
                '   CardResult = DO_ReadPort(card, Channel_P1CH, LightOFF)
                  Call MsecDelay(1)
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                   If CardResult <> 0 Then
                    MsgBox "Read card detect light off fail"
                    End
                   End If
                   
                  
                   
                 
                '===========================================
                 'NO card test
                 '============================================
               '   CardResult = DO_WritePort(card, Channel_P1CL, &HA)  'Power Enable + Slot0 enable
                    CardResult = DO_WritePort(card, Channel_P1A, &H4)
                   If CardResult <> 0 Then
                    MsgBox "Set card detect light on fail"
                    End
                   End If
                   Call MsecDelay(0.6)
                  '  CardResult = DO_ReadPort(card, Channel_P1CH, LightON)
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                   If CardResult <> 0 Then
                    MsgBox "Read light on fail"
                    End
                   End If
                   
                   
                  
                '===========================================
                 'R/W test
                 '============================================
                  
               ' If LightOFF <> 14 Or LightON <> 12 Then
                  Call MsecDelay(0.6)
               If Left(ChipName, 10) <> "AU9520ALF2" Then
                
               
                  If LightOff <> 251 Or LightOn <> 248 Then
                        Tester.Cls
                        Tester.Print "light fail"
                        Tester.Print "LightOFF="; LightOff
                         Tester.Print "LightON="; LightOn
                        
                           AU9520FLSingleSlot = 2
                   
                Else
                       
                         
                        Tester.Cls
                        
                        strReaderName = VenderString_AU9520FLSingleMode & " " & ProductString_AU9520FLSingleMode & " 0"   'slot 0
                        udtReaderStates(0).szReader = strReaderName & vbNullChar
                        Tester.txtmsg.Text = udtReaderStates(0).szReader
                        Tester.Print "Start to test slot0!"
                        Call StartTest
                         AU9520FLSingleSlot = SmartCardTestResult
                 End If
                 
             End If
             
   '=========for   ChipName = "AU9520ALF20"==================================================
  '
   '            If Left(ChipName, 10) = "AU9520ALF2" Then
                
               
   '                If LightOFF <> 255 Or LightON <> 253 Then
               
   '              Tester.Cls
   '              Tester.Print "light fail"
   '              Tester.Print "LightOFF="; LightOFF
   '               Tester.Print "LightON="; LightON
                 
   '                 AU9520FLSingleSlot = 2
                   
                   
                   
    '            Else
                       
                         
     '                   Tester.Cls
                        
      '                  strReaderName = VenderString_AU9520JJ & " " & ProductString_AU9520JJ & " 0"   'slot 0
      '                  udtReaderStates(0).szReader = strReaderName & vbNullChar
      '                  Tester.txtmsg.Text = udtReaderStates(0).szReader
      '                  Tester.Print "Start to test slot0!"
      '                  Call StartTest
       '                 AU9520JJMode = SmartCardTestResult
       '          End If
       '        End If
 End Function
 
 Public Function AU9520FLF21SingleSlot() As Integer
On Error Resume Next
Dim tmpStr As String
tmpStr = ""

'(1) this is for JJ mode  --- one slot mode
     
FindSmartCardResult = 0
SmartCardTestResult = 0
                If PCI7248InitFinish = 0 Then
                      PCI7248Exist
                End If
                
                '3 bit: rom selection
                '2 bit: pwr enable
                '
      
                
                '=========================================
                '    POWER on
                '=========================================
               
                 
                    result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    result = DIO_PortConfig(card, Channel_P1A, OUTPUT_PORT)
               
                   CardResult = DO_WritePort(card, Channel_P1A, &H2)
                  
                  Call MsecDelay(0.4)
                  'CardResult = DO_WritePort(card, Channel_P1CL, &HF)
                  CardResult = DO_WritePort(card, Channel_P1A, &H7)
               
                 CardResult = DO_WritePort(card, Channel_P1A, &H81)  'Power Enable
                 Call MsecDelay(2#)      'power on time, for unkn
                 
                 If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
                 
           
                 
                
                 result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                 '===========================================
                 'NO card test
                 '============================================
                '   CardResult = DO_ReadPort(card, Channel_P1CH, LightOFF)
                 
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                   If CardResult <> 0 Then
                    MsgBox "Read card detect light off fail"
                    End
                    
                   End If
                   
                       tmpStr = GetDeviceName("vid_1b0e")
                 
                   If tmpStr = "" Then
                        AU9520FLF21SingleSlot = 0
                        Exit Function
                 End If
                 
                  
                   
                 
                '===========================================
                 'NO card test
                 '============================================
               '   CardResult = DO_WritePort(card, Channel_P1CL, &HA)  'Power Enable + Slot0 enable
                    CardResult = DO_WritePort(card, Channel_P1A, &H4)
                   If CardResult <> 0 Then
                    MsgBox "Set card detect light on fail"
                    End
                   End If
                   Call MsecDelay(1.6)
                  '  CardResult = DO_ReadPort(card, Channel_P1CH, LightON)
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                   If CardResult <> 0 Then
                    MsgBox "Read light on fail"
                    End
                   End If
                   
                   
                  
                '===========================================
                 'R/W test
                 '============================================
                  
               ' If LightOFF <> 14 Or LightON <> 12 Then
                 
               If Left(ChipName, 10) <> "AU9520ALF2" Then
                
               
                  If LightOff <> 251 Or LightOn <> 248 Then
                        Tester.Cls
                        Tester.Print "light fail"
                        Tester.Print "LightOFF="; LightOff
                         Tester.Print "LightON="; LightOn
                        
                           AU9520FLF21SingleSlot = 2
                   
                Else
                       
                         
                        Tester.Cls
                        
                        strReaderName = VenderString_AU9520FLSingleMode & " " & ProductString_AU9520FLSingleMode & " 0"   'slot 0
                        udtReaderStates(0).szReader = strReaderName & vbNullChar
                        Tester.txtmsg.Text = udtReaderStates(0).szReader
                        Tester.Print "Start to test slot0!"
                        Call StartTest
                         AU9520FLF21SingleSlot = SmartCardTestResult
                 End If
                 
             End If
             
   '=========for   ChipName = "AU9520ALF20"==================================================
  '
   '            If Left(ChipName, 10) = "AU9520ALF2" Then
                
               
   '                If LightOFF <> 255 Or LightON <> 253 Then
               
   '              Tester.Cls
   '              Tester.Print "light fail"
   '              Tester.Print "LightOFF="; LightOFF
   '               Tester.Print "LightON="; LightON
                 
   '                 AU9520FLSingleSlot = 2
                   
                   
                   
    '            Else
                       
                         
     '                   Tester.Cls
                        
      '                  strReaderName = VenderString_AU9520JJ & " " & ProductString_AU9520JJ & " 0"   'slot 0
      '                  udtReaderStates(0).szReader = strReaderName & vbNullChar
      '                  Tester.txtmsg.Text = udtReaderStates(0).szReader
      '                  Tester.Print "Start to test slot0!"
      '                  Call StartTest
       '                 AU9520JJMode = SmartCardTestResult
       '          End If
       '        End If
 End Function
  
  
  Public Function AU9520GLSingleSlot() As Integer
On Error Resume Next
Dim tmpStr As String
tmpStr = ""

'(1) this is for JJ mode  --- one slot mode
     
FindSmartCardResult = 0
SmartCardTestResult = 0
                If PCI7248InitFinish = 0 Then
                      PCI7248Exist
                End If
                
                '3 bit: rom selection
                '2 bit: pwr enable
                '
                
                
               
                
                
                '=========================================
                '    POWER on
                '=========================================
                 ' result = DIO_PortConfig(card, Channel_P1CH, OUTPUT_PORT)
                 ' CardResult = DO_WritePort(card, Channel_P1CH, &H0)
                 '  CardResult = DO_WritePort(card, Channel_P1CL, &H4)  'Power Enable
                 
                    result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    result = DIO_PortConfig(card, Channel_P1A, OUTPUT_PORT)
                    CardResult = DO_WritePort(card, Channel_P1A, &H2)
                  
                  Call MsecDelay(0.3)
                  'CardResult = DO_WritePort(card, Channel_P1CL, &HF)
                  CardResult = DO_WritePort(card, Channel_P1A, &H7)
               '   CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.2)
              '   CardResult = DO_WritePort(card, Channel_P1CL, &HB)  'Power Enable
                 CardResult = DO_WritePort(card, Channel_P1A, &H81)  'Power Enable
                 Call MsecDelay(1.5)    'power on time
                 
                 If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
                 
               tmpStr = GetDeviceName("vid_058f")
                 
                   If tmpStr = "" Then
                        AU9520GLSingleSlot = 0
                        Exit Function
                 End If
                 
                 
                 ' result = DIO_PortConfig(card, Channel_P1CH, INPUT_PORT)
                 result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                 '===========================================
                 'NO card test
                 '============================================
                '   CardResult = DO_ReadPort(card, Channel_P1CH, LightOFF)
                  Call MsecDelay(1)
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                   If CardResult <> 0 Then
                    MsgBox "Read card detect light off fail"
                    End
                   End If
                   
                  
                   
                 
                '===========================================
                 'NO card test
                 '============================================
               '   CardResult = DO_WritePort(card, Channel_P1CL, &HA)  'Power Enable + Slot0 enable
                    CardResult = DO_WritePort(card, Channel_P1A, &H4)
                   If CardResult <> 0 Then
                    MsgBox "Set card detect light on fail"
                    End
                   End If
                   Call MsecDelay(0.6)
                  '  CardResult = DO_ReadPort(card, Channel_P1CH, LightON)
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                   If CardResult <> 0 Then
                    MsgBox "Read light on fail"
                    End
                   End If
                   
                   
                  
                '===========================================
                 'R/W test
                 '============================================
                  
               ' If LightOFF <> 14 Or LightON <> 12 Then
               
               If Left(ChipName, 10) <> "AU9520ALF2" Then
                
               
                  If LightOff <> 251 Or LightOn <> 248 Then
                 Tester.Cls
                 Tester.Print "light fail"
                 Tester.Print "LightOFF="; LightOff
                  Tester.Print "LightON="; LightOn
                 
                    AU9520GLSingleSlot = 2
                   
                   
                  Else
                
                       
                         
                        Tester.Cls
                        
                        strReaderName = VenderString_AU9520GLSingleMode & " " & ProductString_AU9520GLSingleMode & " 0"   'slot 0
                        udtReaderStates(0).szReader = strReaderName & vbNullChar
                        Tester.txtmsg.Text = udtReaderStates(0).szReader
                        Tester.Print "Start to test slot0!"
                        Call StartTest
                         AU9520GLSingleSlot = SmartCardTestResult
                 End If
                 
             End If
             
   '=========for   ChipName = "AU9520ALF20"==================================================
  '
   '            If Left(ChipName, 10) = "AU9520ALF2" Then
                
               
   '                If LightOFF <> 255 Or LightON <> 253 Then
               
   '              Tester.Cls
   '              Tester.Print "light fail"
   '              Tester.Print "LightOFF="; LightOFF
   '               Tester.Print "LightON="; LightON
                 
   '                 AU9520GLSingleSlot = 2
                   
                   
                   
    '            Else
                       
                         
     '                   Tester.Cls
                        
      '                  strReaderName = VenderString_AU9520JJ & " " & ProductString_AU9520JJ & " 0"   'slot 0
      '                  udtReaderStates(0).szReader = strReaderName & vbNullChar
      '                  Tester.txtmsg.Text = udtReaderStates(0).szReader
      '                  Tester.Print "Start to test slot0!"
      '                  Call StartTest
       '                 AU9520JJMode = SmartCardTestResult
       '          End If
       '        End If
 End Function
 
Public Function AU9562GFF20TestSub() As Integer

On Error Resume Next
Dim tmpStr As String
Dim tmpCounter As Integer
Dim mMsg As msg

Dim OldTimer
Dim PassTime
Dim AlcorMessage As Long
    tmpStr = ""
    tmpCounter = 0

    If PCI7248InitFinish = 0 Then
          PCI7248Exist
    End If

    '=========================================
    '    POWER on
    '=========================================
'    CardResult = DO_WritePort(card, Channel_P1A, &H1)       ' for pw on and set usb mode
'    Call MsecDelay(0.1)
    CardResult = DO_WritePort(card, Channel_P1A, &H4)       ' for pw on and set usb mode
    
    If CardResult <> 0 Then
        MsgBox "Power on fail"
        End
    End If
    
    Call MsecDelay(0.2)
    rv0 = WaitDevOn("vid_058f")
    Call MsecDelay(0.3)


    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
    If CardResult <> 0 Then
        MsgBox "Read card detect light off fail"
        End
    End If
           
    If rv0 <> 1 Then
        AU9562GFF20TestSub = 0
        rv0 = 0
        Call LabelMenu(0, rv0, 1)
        GoTo AU9562End_Label
    End If
    
    '===========================================
    'NO card test
    '============================================
    Tester.Cls
                    
    ' must check on PC when replace driver
    strReaderName = VenderString_AU9520GLSingleMode & " " & ProductString_AU9520GLSingleMode & " 0"   'slot 0
    udtReaderStates(0).szReader = strReaderName & vbNullChar
    Tester.txtmsg.Text = udtReaderStates(0).szReader
    Tester.Print "Start to test slot0!"
    Call StartTest
    
    AU9562GFF20TestSub = SmartCardTestResult
    
    rv0 = SmartCardTestResult
    If rv0 = 0 Then
       rv0 = 2
    End If
      
    Call LabelMenu(0, rv0, 1)
    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
    If CardResult <> 0 Then
       MsgBox "Read card detect light on fail"
       End
    End If
                             
    rv1 = 1
    If LightOn <> 252 Then
        Tester.Cls
        Tester.Print "light fail"
        Tester.Print "LightOFF="; LightOff
        rv1 = 2
    End If
    
    Call LabelMenu(1, rv1, rv0)
                             
AU9562End_Label:

    CardResult = DO_WritePort(card, Channel_P1A, &H1)       ' for pw off
    WaitDevOFF ("vid_058f")
    Call MsecDelay(0.2)

    If rv0 * rv1 = 1 Then
         
        CardResult = DO_WritePort(card, Channel_P1A, &H1)       ' for pw on and set usb mode
        Call MsecDelay(0.1)
        CardResult = DO_WritePort(card, Channel_P1A, &H2)       ' for pw on and set rs232 mode
         
        LoadAP_Click_AU9562
        
        OldTimer = Timer
        
        Do
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
            
            PassTime = Timer - OldTimer
            
        Loop Until AlcorMessage = WM_COM_TEST_PASS _
                Or AlcorMessage = WM_COM_TEST_FAIL _
                Or AlcorMessage = WM_COM_TEST_UNKNOWN _
                Or AlcorMessage = WM_COM_TEST_CLOSE _
                Or PassTime > 10
                
        If AlcorMessage = WM_COM_TEST_PASS Then
            TestResult = "PASS"
        ElseIf AlcorMessage = WM_COM_TEST_FAIL Then
            TestResult = "Bin4"                             ' for rs232 fail
        End If
        
    ElseIf rv0 = 0 Then
        TestResult = "Bin2"                                 ' for unknown fail
    Else
        TestResult = "Bin3"                                 ' for usb fail
    End If
    
    CardResult = DO_WritePort(card, Channel_P1A, &H1)       ' for pw on and set usb mode

End Function
 
Public Sub LoadAP_Click_AU9562()
Dim TimePass
Dim rt2
Dim mMsg As msg
Dim winHwnd As Long

    ' find window
    winHwnd = FindWindow(vbNullString, "9540COMAP")
     
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(Tester.hwnd, "open", App.Path & "\AU9562\9540COMAP.exe", "", "", SW_SHOW)
    End If
    
    Call MsecDelay(0.1)
    winHwnd = FindWindow(vbNullString, "9540COMAP")
    
    If winHwnd <> 0 Then
        rt2 = PostMessage(winHwnd, WM_COM_START_TEST, 0&, 0&)
    End If
    

End Sub
 
Public Function AU9562AFE10TestSub() As Integer
On Error Resume Next

Dim OutBuffer(8) As Byte
Dim lngControlReplyLen As Long
Dim InBuffer(1) As Byte
Dim byMaxVal, byMinVal As Byte
Dim lngResult As Long
Dim iCount As Integer

Const CycleLimit = 10
Const Interval = 0.05

OutBuffer(0) = &H40
OutBuffer(1) = &HC6
OutBuffer(2) = &H0
OutBuffer(3) = &HE4
OutBuffer(4) = &H20
OutBuffer(5) = &H1
OutBuffer(6) = &H0
OutBuffer(7) = &H0
FindSmartCardResult = 0

If (lngResult <> 0) Then
    OutMsg "Can't connect to the card"
End If

If PCI7248InitFinish = 0 Then
    PCI7248Exist
End If
 
 
rv0 = 0     'Enum
rv1 = 0     'Cmd fail
rv2 = 1     'Val < 6
rv3 = 0     'MaxVal - MinVal >=2

Tester.Cls
Tester.Print "Begin Test ..."
 
'=========================================
'    POWER on
'=========================================
                  
CardResult = DO_WritePort(card, Channel_P1A, &H0)
'Call MsecDelay(1.6)   ' for unknow device
If CardResult <> 0 Then
   MsgBox "Power on fail"
   End
End If
     
     
Call MsecDelay(0.2)
rv0 = WaitDevOn("vid_058f")
'Call MsecDelay(0.3)


'=========================
'Get MSB tunner value
'=========================
If rv0 = 1 Then

    'Connect to Device
                                        
    strReaderName = VenderString_AU9520GLSingleMode & " " & ProductString_AU9520GLSingleMode & " 0"   'slot 0
    udtReaderStates(0).szReader = strReaderName & vbNullChar
    Tester.txtmsg.Text = udtReaderStates(0).szReader
    
    If (AlcorFindTheCard() = False) Then
        rv0 = 2
    End If
    
    
    If rv0 = 1 Then
        lngPreProtocol = SCARD_PROTOCOL_T0 Or SCARD_PROTOCOL_T1
        lngResult = SCardConnectA(lngContext, strReaderName, SCARD_SHARE_SHARED, lngPreProtocol, lngCard, lngActiveProtocol)
        If (lngResult <> 0) Then
            OutMsg "Can't connect to the card"
            GoTo TestDoneLabel
        End If
        
        
        For iCount = 1 To 10
            Call MsecDelay(Interval)
            
            lngResult = SCardControl(lngCard, IOCTL_SMC_WRITE_READ, OutBuffer(0), 8, _
                                    InBuffer(0), 1, lngControlReplyLen)
            
            If lngResult = 0 Then
                rv1 = 1
                OutMsg "Cycle: " & iCount & " ,Value = 0x" & Byte2Char(InBuffer(0))
'                If InBuffer(0) >= 6 Then
'                    rv2 = 1
                    
                If iCount = 1 Then
                    byMaxVal = (InBuffer(0) And &HF)
                    byMinVal = (InBuffer(0) And &HF)
                End If
                    
'                Else
'                    rv2 = 0
'                    Exit For
'                End If
                
            Else
                rv1 = 0
                Exit For
            End If
            
            
            If byMaxVal < (InBuffer(0) And &HF) Then
                byMaxVal = (InBuffer(0) And &HF)
            End If
            
            If byMinVal > (InBuffer(0) And &HF) Then
                byMinVal = (InBuffer(0) And &HF)
            End If
            
            
        Next
    End If
End If


If (rv2 = 1) Then
    Tester.Print "Max = " & byMaxVal & vbCrLf & "Min = " & byMinVal

    If (Abs(byMaxVal - byMinVal) < 2) Then
        rv3 = 1
    Else
        rv3 = 0
    End If
End If


TestDoneLabel:

CardResult = DO_WritePort(card, Channel_P1A, &HFF)
WaitDevOFF ("vid_058f")
Call MsecDelay(0.2)

If (rv0 <> 1) Then
    TestResult = "Bin2"
ElseIf (rv1 <> 1) Then
    TestResult = "Bin3"
ElseIf (rv2 <> 1) Then
    TestResult = "Bin4"
ElseIf (rv3 <> 1) Then
    TestResult = "Bin5"
ElseIf (rv0 * rv1 * rv2 * rv3 = 1) Then

    TestResult = "PASS"
End If


End Function

Public Function AU9520ASF20TestSub() As Integer
On Error Resume Next
Dim tmpStr As String
Dim tmpCounter As Integer
tmpStr = ""
tmpCounter = 0

'(1) this is for JJ mode  --- one slot mode
     
FindSmartCardResult = 0
SmartCardTestResult = 0
                If PCI7248InitFinish = 0 Then
                      PCI7248Exist
                End If
                
             
               Call MsecDelay(0.5)
                
                
                '=========================================
                '    POWER on
                '=========================================
                                 
                  CardResult = DO_WritePort(card, Channel_P1A, &H0)
                  'Call MsecDelay(1.6)   ' for unknow device
                 If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
                    
                    
                 Call MsecDelay(0.2)
                 rv0 = WaitDevOn("vid_058f")
                 Call MsecDelay(0.3)
                 
                 'NO card test
                 '============================================
                 LightOff = 0
                 
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                   If CardResult <> 0 Then
                    MsgBox "Read card detect light off fail"
                    End
                   End If
                   
                  
                   '    tmpStr = GetDeviceName("vid_058f")
                 
                   ' rv0 = 2
                   If rv0 <> 1 Then
                        AU9520ASF20TestSub = 0
                        rv0 = 0
                         Call LabelMenu(0, rv0, 1)
                        GoTo AU9520ASF20Result
                 End If
                 
          
                 '===========================================
                 'NO card test
                 '============================================
                '   CardResult = DO_ReadPort(card, Channel_P1CH, LightOFF)
                '  Call MsecDelay(1)
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                   If CardResult <> 0 Then
                    MsgBox "Read card detect light off fail"
                    End
                   End If
                   
                  
                   
     
                         
                        Tester.Cls
                                         
                        strReaderName = VenderString_AU9520GLSingleMode & " " & ProductString_AU9520GLSingleMode & " 0"   'slot 0
                        udtReaderStates(0).szReader = strReaderName & vbNullChar
                        Tester.txtmsg.Text = udtReaderStates(0).szReader
                        Tester.Print "Start to test slot0!"
                        Call StartTest
                         AU9520ASF20TestSub = SmartCardTestResult
                         
                           rv0 = AU9520ASF20TestSub
                           If rv0 = 0 Then
                              rv0 = 2
                           End If
                           
                         Call LabelMenu(0, rv0, 1)
                          CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If CardResult <> 0 Then
                            MsgBox "Read card detect light off fail"
                            End
                         End If
                         
                            rv1 = 1
                             If LightOff <> 252 Then
                                Tester.Cls
                                Tester.Print "light fail"
                                Tester.Print "LightOFF="; LightOff
                               rv1 = 2
                             End If
                        Call LabelMenu(1, rv1, rv0)
                         
AU9520ASF20Result:
                
                CardResult = DO_WritePort(card, Channel_P1A, &HFF)
'                Call MsecDelay(0.4)
'
'                Do
'                    tmpCounter = tmpCounter + 1
'                    tmpStr = GetDeviceName("vid_058f")
'                    Call MsecDelay(0.2)
'                Loop Until ((tmpStr <> "") Or (tmpCounter >= 5))

                WaitDevOFF ("vid_058f")
                Call MsecDelay(0.3)
                
                
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv1 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv1 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
                    TestResult = "CF_RF"
                ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                 ElseIf rv3 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv3 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv0 * rv1 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
    
 End Function
 
Public Function AU9562BSF40TestSub() As Integer
On Error Resume Next
Dim tmpStr As String
Dim tmpCounter As Integer

Dim OutBuffer(8) As Byte
Dim lngControlReplyLen As Long
Dim InBuffer(1) As Byte
Dim lngResult As Long

tmpStr = ""
tmpCounter = 0

'(1) this is for JJ mode  --- one slot mode
     
FindSmartCardResult = 0
SmartCardTestResult = 0

If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If


Call MsecDelay(0.5)
                
                
'=========================================
'    POWER on
'=========================================
                
CardResult = DO_WritePort(card, Channel_P1A, &H0)
'Call MsecDelay(1.6)   ' for unknow device
If CardResult <> 0 Then
   MsgBox "Power on fail"
   End
End If
   
Call MsecDelay(0.2)
rv0 = WaitDevOn("vid_058f")
Call MsecDelay(0.3)
                 
'NO card test
'============================================
LightOff = 0

CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
If CardResult <> 0 Then
    MsgBox "Read card detect light off fail"
    End
End If
                   
                  
'tmpStr = GetDeviceName("vid_058f")

'rv0 = 2
If rv0 <> 1 Then
    AU9562BSF40TestSub = 0
    rv0 = 0
    Call LabelMenu(0, rv0, 1)
    GoTo AU9520ASF20Result
End If
                 
          
'===========================================
'NO card test
'============================================
'CardResult = DO_ReadPort(card, Channel_P1CH, LightOFF)
'Call MsecDelay(1)
CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
If CardResult <> 0 Then
    MsgBox "Read card detect light off fail"
    End
End If
              
                         
Tester.Cls
                
strReaderName = VenderString_AU9520GLSingleMode & " " & ProductString_AU9520GLSingleMode & " 0"   'slot 0
udtReaderStates(0).szReader = strReaderName & vbNullChar
Tester.txtmsg.Text = udtReaderStates(0).szReader
Tester.Print "Start to test slot0!"
Call StartTest
                                                  
rv0 = SmartCardTestResult
If rv0 = 0 Then
    rv0 = 2
End If
                           
Call LabelMenu(0, rv0, 1)
CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
If CardResult <> 0 Then
   MsgBox "Read card detect light off fail"
   End
End If

rv1 = 1
If LightOff <> 252 Then
    Tester.Cls
    Tester.Print "light fail"
    Tester.Print "LightOFF="; LightOff
    rv1 = 2
End If
Call LabelMenu(1, rv1, rv0)

'20140814 for AU9562 marking to AU9540, purpose to avoid xtal bonding fail
                        
If (rv0 * rv1 = 1) Then
    Tester.txtmsg.Text = ""
    lngPreProtocol = SCARD_PROTOCOL_T0 Or SCARD_PROTOCOL_T1
    lngResult = SCardConnectA(lngContext, strReaderName, SCARD_SHARE_SHARED, lngPreProtocol, lngCard, lngActiveProtocol)
    If (lngResult <> 0) Then
        OutMsg "Can't connect to the card"
        rv0 = 2
        Call LabelMenu(1, rv0, rv1)
    End If
End If

If (rv0 * rv1 = 1) Then

    rv0 = 0
    OutBuffer(0) = &H40
    OutBuffer(1) = &HC6
    OutBuffer(2) = &H0
    OutBuffer(3) = &HE2
    OutBuffer(4) = &H20
    OutBuffer(5) = &H1
    OutBuffer(6) = &H0
    OutBuffer(7) = &H0
    
    lngResult = SCardControl(lngCard, IOCTL_SMC_WRITE_READ, OutBuffer(0), 8, _
                             InBuffer(0), 1, lngControlReplyLen)

    If ((lngResult = 0) And (InBuffer(0)) = &H10) Then
        rv0 = 1
        OutMsg "Check XTAL Mode PASS"
    Else
        rv0 = 2
        OutMsg "Check XTAL Mode Fail"
    End If
    Call LabelMenu(1, rv0, rv1)
    
End If

                         
AU9520ASF20Result:
                
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)

    WaitDevOFF ("vid_058f")
    Call MsecDelay(0.3)
                
                
    If rv0 = UNKNOW Then
        UnknowDeviceFail = UnknowDeviceFail + 1
        TestResult = "UNKNOW"
    ElseIf rv0 = WRITE_FAIL Then
        SDWriteFail = SDWriteFail + 1
        TestResult = "SD_WF"
    ElseIf rv0 = READ_FAIL Then
        SDReadFail = SDReadFail + 1
        TestResult = "SD_RF"
    ElseIf rv1 = WRITE_FAIL Then
        CFWriteFail = CFWriteFail + 1
        TestResult = "CF_WF"
    ElseIf rv1 = READ_FAIL Then
        CFReadFail = CFReadFail + 1
        TestResult = "CF_RF"
    ElseIf rv2 = WRITE_FAIL Then
        XDWriteFail = XDWriteFail + 1
        TestResult = "XD_WF"
    ElseIf rv2 = READ_FAIL Then
        XDReadFail = XDReadFail + 1
        TestResult = "XD_RF"
    ElseIf rv0 * rv1 = 1 Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
    
 End Function
 
Public Function AU9562BSF30TestSub() As Integer
On Error Resume Next
Dim tmpStr As String
Dim tmpCounter As Integer
tmpStr = ""
tmpCounter = 0

rv0 = 0 'Enum
rv1 = 0 'Function
rv2 = 0 'LED
rv3 = 0 'OpenShort

'(1) this is for JJ mode  --- one slot mode
     
FindSmartCardResult = 0
SmartCardTestResult = 0

If PCI7248InitFinish = 0 Then
    PCI7248ExistAU6254
    Call SetTimer_1ms
End If

'==========================================================
'    Start Open Shot test
'==========================================================
OS_Result = 0
rv2 = 0

CardResult = DO_WritePort(card, Channel_P1C, &H0)
'
MsecDelay (0.2)

OpenShortTest_SkipZero_NoGPIB

rv3 = OS_Result

If rv3 <> 1 Then
    CardResult = DO_WritePort(card, Channel_P1C, &HFF)
    MsecDelay (0.1)
    CardResult = DO_WritePort(card, Channel_P1C, &H0)
    MsecDelay (0.1)
    OpenShortTest_SkipZero_NoGPIB
    rv3 = OS_Result
End If



CardResult = DO_WritePort(card, Channel_P1C, &HFF)
If rv3 <> 1 Then  'OS Fail
    GoTo AU9562BSF30Result
End If

'=========================================
'    POWER on
'=========================================
'MsecDelay (0.2)
CardResult = DO_WritePort(card, Channel_P1A, &H0)
'Call MsecDelay(1.6)   ' for unknow device
If CardResult <> 0 Then
    MsgBox "Power on fail"
    End
End If


Call MsecDelay(0.2)
rv0 = WaitDevOn("vid_058f")
Call MsecDelay(0.3)

'NO card test
'============================================
LightOff = 0

If rv0 <> 1 Then
    rv0 = 0
    Call LabelMenu(0, rv0, 1)
    GoTo AU9562BSF30Result
End If


'Tester.Cls
                 
strReaderName = VenderString_AU9520GLSingleMode & " " & ProductString_AU9520GLSingleMode & " 0"   'slot 0
udtReaderStates(0).szReader = strReaderName & vbNullChar
Tester.txtmsg.Text = udtReaderStates(0).szReader
Tester.Print "Start to test slot0!"

Call StartTest
rv1 = SmartCardTestResult

If rv1 = 0 Then
   rv1 = 2
End If
   
Call LabelMenu(0, rv1, 1)
CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
If CardResult <> 0 Then
   MsgBox "Read card detect light off fail"
   End
End If

If rv1 = 1 Then
    rv2 = 1
    If LightOff <> 252 Then
        Tester.Cls
        Tester.Print "light fail"
        Tester.Print "LightOFF="; LightOff
        rv2 = 2
    End If
    Call LabelMenu(1, rv2, rv1)
End If

                         
                         
                         
AU9562BSF30Result:
                
    CardResult = DO_WritePort(card, Channel_P1A, &H80)
    WaitDevOFF ("vid_058f")
    Call MsecDelay(0.3)
                
    If rv3 <> 1 Then        'OpenShort
        TestResult = "Bin2"
    ElseIf rv0 <> 1 Then    'Enum
       UnknowDeviceFail = UnknowDeviceFail + 1
       TestResult = "Bin4"
    ElseIf rv1 <> 1 Then    'Function
        TestResult = "Bin3"
    ElseIf rv2 <> 1 Then    'LED
        TestResult = "Bin5"
    ElseIf rv0 * rv1 * rv2 * rv3 = 1 Then
         TestResult = "PASS"
    Else
        TestResult = "Bin2"
      
    End If
    
End Function
 
Public Function AU9540BSF00TestSub() As Integer

On Error Resume Next
Dim tmpStr As String
Dim tmpCounter As Integer
Dim LV_Flag As Boolean
Dim HV_Result As String
Dim LV_Result As String

    tmpStr = ""
    tmpCounter = 0

    '(1) this is for JJ mode  --- one slot mode
    
    FindSmartCardResult = 0
    SmartCardTestResult = 0
    
    If PCI7248InitFinish = 0 Then
        PCI7248Exist
    End If


    '=========================================
    '    POWER on
    '=========================================
    
    LV_Flag = False
    HV_Result = ""
    LV_Result = ""
    Tester.Cls

Routine_Label:
                 
    If Not LV_Flag Then
        Call PowerSet2(0, "3.6", "0.5", 1, "3.6", "0.5", 1)
        Tester.Print "AU9540BS : HV Begin Test ..."
    Else
        Call PowerSet2(0, "3.0", "0.5", 1, "3.0", "0.5", 1)
        Tester.Print "AU9540BS : LV Begin Test ..."
    End If
    
    CardResult = DO_WritePort(card, Channel_P1A, &H0)
    'Call MsecDelay(1.6)   ' for unknow device
    If CardResult <> 0 Then
        MsgBox "Power on fail"
        End
    End If

    Call MsecDelay(0.5)
    rv0 = WaitDevOn("vid_058f")
    Call MsecDelay(0.3)

    'NO card test
    '============================================
    
    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
    If CardResult <> 0 Then
        MsgBox "Read card detect light off fail"
        End
    End If

    If rv0 <> 1 Then
        AU9540BSF00TestSub = 0
        rv0 = 0
        Call LabelMenu(0, rv0, 1)
        GoTo AU9540BSF00Result
    End If


    '===========================================
    'NO card test
    '============================================

    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
    If CardResult <> 0 Then
        MsgBox "Read card detect light off fail"
        End
    End If

    Tester.Cls
    
    strReaderName = VenderString_AU9520GLSingleMode & " " & ProductString_AU9520GLSingleMode & " 0"   'slot 0
    udtReaderStates(0).szReader = strReaderName & vbNullChar
    Tester.txtmsg.Text = udtReaderStates(0).szReader
    Tester.Print "Start to test slot0!"
    Call StartTest
    AU9540BSF00TestSub = SmartCardTestResult

    rv0 = AU9540BSF00TestSub
    If rv0 = 0 Then
        rv0 = 2
    End If

    Call LabelMenu(0, rv0, 1)
    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
    If CardResult <> 0 Then
        MsgBox "Read card detect light off fail"
        End
    End If

    rv1 = 1
    If LightOff <> 252 Then
        Tester.Cls
        Tester.Print "light fail"
        Tester.Print "LightOFF="; LightOff
        rv1 = 2
    End If
    Call LabelMenu(1, rv1, rv0)

AU9540BSF00Result:

    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
    Call PowerSet2(0, "0.0", "0.5", 1, "0.0", "0.5", 1)
    Call MsecDelay(0.4)
    WaitDevOFF ("058f")
    
    Do
        tmpCounter = tmpCounter + 1
        tmpStr = GetDeviceName("vid_058f")
        Call MsecDelay(0.2)
    Loop Until ((tmpStr <> "") Or (tmpCounter >= 5))

    If Not LV_Flag Then
        If rv0 = 0 Then
            LV_Result = "Bin2"
        ElseIf rv0 * rv1 <> 1 Then
            LV_Result = "Fail"
        ElseIf rv0 * rv1 = 1 Then
            LV_Result = "PASS"
        Else
            LV_Result = "Fail"
        End If
        
        rv0 = 0
        rv1 = 0
        LV_Flag = True
        GoTo Routine_Label
    Else
        If rv0 = 0 Then
            HV_Result = "Bin2"
        ElseIf rv0 * rv1 <> 1 Then
            HV_Result = "Fail"
        ElseIf rv0 * rv1 = 1 Then
            HV_Result = "PASS"
        Else
            HV_Result = "Fail"
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
    
 End Function
 
 
Public Function AU9520ASF00TestSub() As Integer
On Error Resume Next
Dim tmpStr As String
Dim tmpCounter As Integer
Dim HV_TestDoneFlag As Boolean
Dim HV_Result As String
Dim LV_Result As String
tmpStr = ""
tmpCounter = 0


HV_Result = 0
LV_Result = 0

'(1) this is for JJ mode  --- one slot mode
     
FindSmartCardResult = 0
SmartCardTestResult = 0

HV_TestDoneFlag = False

If PCI7248InitFinish_Sync = 0 Then
    PCI7248Exist_P1C_Sync
End If
                
Tester.Cls

Routine_Label:


rv0 = 0 'Enum
rv1 = 0 'R/W
rv2 = 0 'Led

SetSiteStatus (SiteReady)
Call WaitAnotherSiteDone(SiteReady, 3#)

'=========================================
'    POWER on
'=========================================
If Not HV_TestDoneFlag Then
    Tester.Print "Begin HV(3.6) Test ..."
    Call PowerSet2(0, "3.6", "0.5", 1, "3.6", "0.5", 1)
Else
    Tester.Print "Begin LV(3.0) Test ..."
    Call PowerSet2(0, "3.0", "0.5", 1, "3.0", "0.5", 1)
End If
                
CardResult = DO_WritePort(card, Channel_P1A, &H0)
'Call MsecDelay(1.6)   ' for unknow device
If CardResult <> 0 Then
    MsgBox "Power on fail"
    End
End If
   
Call MsecDelay(0.3)
rv0 = WaitDevOn("vid_058f")
Call MsecDelay(0.5)

'===========================================
'NO card test
'============================================
CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
If CardResult <> 0 Then
    MsgBox "Read card detect light off fail"
    End
End If

Call NewLabelMenu(0, "UnknowDevice", rv0, 1)

If rv0 = 1 Then
    strReaderName = VenderString_AU9520GLSingleMode & " " & ProductString_AU9520GLSingleMode & " 0"   'slot 0
    udtReaderStates(0).szReader = strReaderName & vbNullChar
    Tester.txtmsg.Text = udtReaderStates(0).szReader
    Tester.Print "Start to test slot0!"
    Call StartTest
    rv1 = SmartCardTestResult
    Call NewLabelMenu(1, "RW Test", rv1, rv0)
End If


If rv1 = 1 Then
    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)

    rv2 = 1
    If LightOff <> 252 Then
        Tester.Cls
        Tester.Print "light fail"
        Tester.Print "LightOFF="; LightOff
        rv2 = 2
    End If
    Call NewLabelMenu(2, "Led Fail", rv2, rv1)
End If


AU9520ASF00Result:

SetSiteStatus (HVDone)
Call WaitAnotherSiteDone(HVDone, 3#)

CardResult = DO_WritePort(card, Channel_P1A, &HFF)
Call PowerSet2(0, "0.0", "0.5", 1, "0.0", "0.5", 1)
Call MsecDelay(0.2)
WaitDevOFF ("058f")


If Not HV_TestDoneFlag Then
    If rv0 = 0 Then
        HV_Result = "Unknow"
    ElseIf rv0 * rv1 * rv2 <> 1 Then
        HV_Result = "Fail"
    ElseIf rv0 * rv1 * rv2 = 1 Then
        HV_Result = "PASS"
    End If
    
    Tester.Print "HV Test " & HV_Result & vbCrLf
    HV_TestDoneFlag = True
    GoTo Routine_Label
Else
    If rv0 = 0 Then
        LV_Result = "Unknow"
    ElseIf rv0 * rv1 * rv2 <> 1 Then
        LV_Result = "Fail"
    ElseIf rv0 * rv1 * rv2 = 1 Then
        LV_Result = "PASS"
    End If
    
    Tester.Print "LV Test " & LV_Result
End If



If (HV_Result = "Unknow") And (LV_Result = "Unknow") Then
    TestResult = "Bin2"
ElseIf (HV_Result = "Fail") And (LV_Result = "PASS") Then
    TestResult = "Bin3"
ElseIf (HV_Result = "PASS") And (LV_Result = "Fail") Then
    TestResult = "Bin4"
ElseIf (HV_Result = "Fail") And (LV_Result = "Fail") Then
    TestResult = "Bin5"
ElseIf (HV_Result = "PASS") And (LV_Result = "PASS") Then
    TestResult = "PASS"
Else
    TestResult = "Bin2"
End If

SetSiteStatus (SiteUnknow)
    
End Function
 
Public Function AU9560BSF20TestSub() As Integer
On Error Resume Next

' 2012/5/16 This Code same as AU9520ASF20TestSub

Dim tmpStr As String
Dim tmpCounter As Integer
tmpStr = ""
tmpCounter = 0

'(1) this is for JJ mode  --- one slot mode
     
FindSmartCardResult = 0
SmartCardTestResult = 0

If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
'=========================================
'    POWER on
'=========================================
                 
CardResult = DO_WritePort(card, Channel_P1A, &H0)

'CardResult = DO_WritePort(card, Channel_P1A, &H4)

'Call MsecDelay(1.6)   ' for unknow device

If CardResult <> 0 Then
    MsgBox "Power on fail"
    End
End If
                    
Call MsecDelay(0.2)
rv0 = WaitDevOn("vid_058f")
Call MsecDelay(0.3)
                 
'NO card test
'============================================

CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
If CardResult <> 0 Then
    MsgBox "Read card detect light off fail"
    End
End If
                   
                  
                   
If rv0 <> 1 Then
    AU9560BSF20TestSub = 0
    rv0 = 0
    Call LabelMenu(0, rv0, 1)
    GoTo AU9560End_Label
End If

          
'===========================================
'NO card test
'============================================
Tester.Cls
                
strReaderName = VenderString_AU9520GLSingleMode & " " & ProductString_AU9520GLSingleMode & " 0"   'slot 0
udtReaderStates(0).szReader = strReaderName & vbNullChar
Tester.txtmsg.Text = udtReaderStates(0).szReader
Tester.Print "Start to test slot0!"
Call StartTest

AU9560BSF20TestSub = SmartCardTestResult

rv0 = SmartCardTestResult
If rv0 = 0 Then
   rv0 = 2
End If
  
Call LabelMenu(0, rv0, 1)
CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
If CardResult <> 0 Then
   MsgBox "Read card detect light on fail"
   End
End If
                         
rv1 = 1
If LightOn <> 252 Then
    Tester.Cls
    Tester.Print "light fail"
    Tester.Print "LightOFF="; LightOff
    rv1 = 2
End If

Call LabelMenu(1, rv1, rv0)
                         
AU9560End_Label:
                
CardResult = DO_WritePort(card, Channel_P1A, &HFF)
WaitDevOFF ("vid_058f")

If rv0 = UNKNOW Then
   UnknowDeviceFail = UnknowDeviceFail + 1
   TestResult = "UNKNOW"
ElseIf rv0 = WRITE_FAIL Then
    SDWriteFail = SDWriteFail + 1
    TestResult = "SD_WF"
ElseIf rv0 = READ_FAIL Then
    SDReadFail = SDReadFail + 1
    TestResult = "SD_RF"
ElseIf rv1 = WRITE_FAIL Then
    CFWriteFail = CFWriteFail + 1
    TestResult = "CF_WF"
ElseIf rv1 = READ_FAIL Then
    CFReadFail = CFReadFail + 1
    TestResult = "CF_RF"
ElseIf rv2 = WRITE_FAIL Then
    XDWriteFail = XDWriteFail + 1
    TestResult = "XD_WF"
ElseIf rv2 = READ_FAIL Then
    XDReadFail = XDReadFail + 1
    TestResult = "XD_RF"
ElseIf rv0 * rv1 = 1 Then
     TestResult = "PASS"
Else
    TestResult = "Bin2"
End If
    
End Function

Public Function AU9560BSF00TestSub() As Integer
On Error Resume Next

' 2012/5/16 This Code same as AU9520ASF20TestSub

Dim tmpStr As String
Dim tmpCounter As Integer
Dim LV_Flag As Boolean
Dim HV_Result As String
Dim LV_Result As String
tmpStr = ""
tmpCounter = 0

'(1) this is for JJ mode  --- one slot mode
     
FindSmartCardResult = 0
SmartCardTestResult = 0

If PCI7248InitFinish = 0 Then
    PCI7248Exist
End If
                
'=========================================
'    POWER on
'=========================================
LV_Flag = False
HV_Result = ""
LV_Result = ""
Tester.Cls

Routine_Label:
                 
If Not LV_Flag Then
    Call PowerSet2(0, "3.6", "0.5", 1, "3.6", "0.5", 1)
    Tester.Print "AU9560BS : HV Begin Test ..."
Else
    Call PowerSet2(0, "3.0", "0.5", 1, "3.0", "0.5", 1)
    Tester.Print "AU9560BS : LV Begin Test ..."
End If

CardResult = DO_WritePort(card, Channel_P1A, &H0)
If CardResult <> 0 Then
    MsgBox "Power on fail"
    End
End If
                    
Call MsecDelay(0.2)
rv0 = WaitDevOn("vid_058f")
Call MsecDelay(0.1)
                 
'NO card test
'============================================

CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
If CardResult <> 0 Then
    MsgBox "Read card detect light off fail"
    End
End If
                   
                  
                   
If rv0 <> 1 Then
    AU9560BSF00TestSub = 0
    rv0 = 0
    Call LabelMenu(0, rv0, 1)
    GoTo AU9560End_Label
End If

          
'===========================================
'NO card test
'============================================
'Tester.Cls
                
strReaderName = VenderString_AU9520GLSingleMode & " " & ProductString_AU9520GLSingleMode & " 0"   'slot 0
udtReaderStates(0).szReader = strReaderName & vbNullChar
Tester.txtmsg.Text = udtReaderStates(0).szReader
Tester.Print "Start to test slot0!"
Call StartTest

AU9560BSF00TestSub = SmartCardTestResult

rv0 = SmartCardTestResult
If rv0 = 0 Then
   rv0 = 2
End If
  
Call LabelMenu(0, rv0, 1)
CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
If CardResult <> 0 Then
   MsgBox "Read card detect light on fail"
   End
End If
                         
rv1 = 1
If LightOn <> 252 Then
    Tester.Print "light fail"
    Tester.Print "LightOFF="; LightOff
    rv1 = 2
End If

Call LabelMenu(1, rv1, rv0)
                         
AU9560End_Label:
                
CardResult = DO_WritePort(card, Channel_P1A, &HFF)
Call PowerSet2(0, "0.0", "0.5", 1, "0.0", "0.5", 1)
WaitDevOFF ("vid_058f")
                
            
If Not LV_Flag Then
    
    If rv0 = 0 Then
        LV_Result = "Bin2"
    ElseIf rv0 * rv1 <> 1 Then
        LV_Result = "Fail"
    ElseIf rv0 * rv1 = 1 Then
        LV_Result = "PASS"
    Else
        LV_Result = "Fail"
    End If
    
    rv0 = 0
    rv1 = 0
    LV_Flag = True
    GoTo Routine_Label
Else
    If rv0 = 0 Then
        HV_Result = "Bin2"
    ElseIf rv0 * rv1 <> 1 Then
        HV_Result = "Fail"
    ElseIf rv0 * rv1 = 1 Then
        HV_Result = "PASS"
    Else
        HV_Result = "Fail"
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
                    
End Function

Public Function AU9562BSF00TestSub() As Integer
On Error Resume Next

'2014/7/1 This Code same as AU9560BSF00

Dim tmpStr As String
Dim tmpCounter As Integer
Dim HV_Done_Flag As Boolean
Dim HV_Result As String
Dim LV_Result As String
tmpStr = ""
tmpCounter = 0

'(1) this is for JJ mode  --- one slot mode
     


If PCI7248InitFinish_Sync = 0 Then
    PCI7248Exist_P1C_Sync
End If
                
'=========================================
'    POWER on
'=========================================


HV_Done_Flag = False
HV_Result = ""
LV_Result = ""
Tester.Cls


Routine_Label:


FindSmartCardResult = 0
SmartCardTestResult = 0


If Not HV_Done_Flag Then
    Call PowerSet2(0, "3.6", "0.5", 1, "3.6", "0.5", 1)
    Tester.Print "AU9562BS : HV Begin Test ..."
    SetSiteStatus (RunHV)
Else
    Call PowerSet2(0, "3.0", "0.5", 1, "3.0", "0.5", 1)
Tester.Print "AU9562BS : LV Begin Test ..."
    SetSiteStatus (RunLV)
End If
Call MsecDelay(0.2)


CardResult = DO_WritePort(card, Channel_P1A, &H0)
If CardResult <> 0 Then
    MsgBox "Power on fail"
    End
End If
                    
Call MsecDelay(0.2)
rv0 = WaitDevOn("vid_058f")
Call MsecDelay(0.1)
                 
'NO card test
'============================================

CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
If CardResult <> 0 Then
    MsgBox "Read card detect light off fail"
    End
End If
                   
                  
                   
If rv0 <> 1 Then
    rv0 = 0
    Call LabelMenu(0, rv0, 1)
    GoTo AU9562End_Label
End If

          
'===========================================
'NO card test
'============================================
'Tester.Cls
                
strReaderName = VenderString_AU9520GLSingleMode & " " & ProductString_AU9520GLSingleMode & " 0"   'slot 0
udtReaderStates(0).szReader = strReaderName & vbNullChar
Tester.txtmsg.Text = udtReaderStates(0).szReader
Tester.Print "Start to test slot0!"
Call StartTest

rv1 = SmartCardTestResult
If rv1 = 0 Then
   rv1 = 2
   GoTo AU9562End_Label
End If
  
Call LabelMenu(0, rv1, 1)
CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
If CardResult <> 0 Then
   MsgBox "Read card detect light on fail"
   End
End If
                         
If rv1 = 1 Then
    If LightOn <> 252 Then
        Tester.Print "light fail"
        Tester.Print "LightOFF="; LightOff
        rv1 = 3
    End If
End If

Call LabelMenu(1, rv1, rv0)
                         
AU9562End_Label:


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
WaitDevOFF ("vid_058f")
SetSiteStatus (SiteUnknow)
    
If HV_Done_Flag = False Then
    If rv0 <> 1 Then
        HV_Result = "Bin2"
        Tester.Print "HV Unknow"
    ElseIf rv0 * rv1 <> 1 Then
        HV_Result = "Fail"
        Tester.Print "HV Fail"
    ElseIf rv0 * rv1 = 1 Then
        HV_Result = "PASS"
        Tester.Print "HV PASS"
    End If
    HV_Done_Flag = True
    Call MsecDelay(0.2)
    ReaderExist = 0
    GoTo Routine_Label
Else
    If rv0 <> 1 Then
        LV_Result = "Bin2"
        Tester.Print "LV Unknow"
    ElseIf rv0 * rv1 <> 1 Then
        LV_Result = "Fail"
        Tester.Print "LV Fail"
    ElseIf rv0 * rv1 = 1 Then
        LV_Result = "PASS"
        Tester.Print "LV PASS"
    End If
    
End If
            
            
If (HV_Result = "Bin2") And (LV_Result = "Bin2") Then
    UnknowDeviceFail = UnknowDeviceFail + 1
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

            
End Function
                        
 Public Function AU9520_2Slot() As Integer
                
 On Error Resume Next
 Dim tmpStr As String
               
FindSmartCardResult = 0
SmartCardTestResult = 0
                
                  If PCI7248InitFinish = 0 Then
                      PCI7248Exist
                End If
                
 '(2) this is for AU9520 mode-- 2 slot
                
                
                
                 '=========================================
                '    POWER on
                '=========================================
               
                 '==========  For old board
                 '   result = DIO_PortConfig(card, Channel_P1CH, OUTPUT_PORT) '--- old
                 '   CardResult = DO_WritePort(card, Channel_P1CH, &H0) '--- old
                 '   CardResult = DO_WritePort(card, Channel_P1CL, &H4) '--- old
                
                      
                      result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    result = DIO_PortConfig(card, Channel_P1A, OUTPUT_PORT)
                     CardResult = DO_WritePort(card, Channel_P1A, &H2)
                              
                 
                  Call MsecDelay(0.4)
                 
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
               '  CardResult = DO_WritePort(card, Channel_P1CL, &H7) --- old
                  CardResult = DO_WritePort(card, Channel_P1A, &H3)  '--- New
                 Call MsecDelay(0.1)
               '   CardResult = DO_WritePort(card, Channel_P1CL, &H3)  'Power Enable                 Call MsecDelay(1)    'power on time
                  CardResult = DO_WritePort(card, Channel_P1A, &H1)
                  Call MsecDelay(2#)  ' for unknow device
                 If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
                 
                
                 
               
                 result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                 '===========================================
                 'NO card test
                 '============================================
                 
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                   If CardResult <> 0 Then
                    MsgBox "Read card detect light off fail"
                    End
                   End If
                   
                  
                       tmpStr = GetDeviceName("vid_058f")
                 
                 If tmpStr = "" Then
                      AU9520_2Slot = 0
                      Exit Function
                 End If
                 
                '===========================================
                 'NO card test
                 '============================================
               '   CardResult = DO_WritePort(card, Channel_P1CL, &H0)  'Power Enable + Slot0 enable
                  CardResult = DO_WritePort(card, Channel_P1A, &H0)
                   If CardResult <> 0 Then
                    MsgBox "Set card detect light on fail"
                    End
                   End If
                   Call MsecDelay(0.6)
                 '   CardResult = DO_ReadPort(card, Channel_P1CH, LightON)
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                   If CardResult <> 0 Then
                    MsgBox "Read light on fail"
                    End
                   End If
                   
                   
                  
                '===========================================
                 'R/W test
                 '============================================
                  
              '  If LightOFF <> 11 Or LightON <> 8 Then
                If LightOff <> 251 Or LightOn <> 248 Then
                     AU9520_2Slot = 2
                     Tester.Print "GPO Fail"
                Else
              
                      
                       
                        strReaderName = VenderString & " " & ProductString & " 0"   'slot 0
                        udtReaderStates(0).szReader = strReaderName & vbNullChar
                        Tester.txtmsg.Text = udtReaderStates(0).szReader
                        
                        Tester.Print "Start to test slot0 !"
                        Call StartTest
                         
                        If SmartCardTestResult = 1 Then
                        
                            SmartCardTestResult = 0  'Reset Test Result
                            Tester.Print "slot0 test ok!"
                            
                            strReaderName = VenderString & " " & ProductString & " 1"   'slot 1
                            udtReaderStates(0).szReader = strReaderName & vbNullChar
                            Tester.txtmsg.Text = udtReaderStates(0).szReader
                            
                           
                            Call StartTest
                            Tester.Print "slot1 test ok!"
                        
                        End If
                        
                        AU9520_2Slot = SmartCardTestResult
               
            End If
                
                
                
                
End Function

Public Sub StartTest()
On Error Resume Next
    ' length of list of readers
    Dim lngReadersLen As Long
    Dim lngReaderstatelen As Long
    ' array of readerstate with only one element
    Dim lngResult As Long
       ' cardinality of readerstate array
    Dim lngReaderNameLen As Long
    Dim lngSCardStatus As Long
    Dim ATR As ARRAY_TYPE
    Dim ATRLen As Long
    Dim i As Long
    
    bStop = False
    
    'Do
   
        lngReaderstatelen = 1
        udtReaderStates(0).dwCurrentState = SCARD_STATE_UNAWARE
        ' check if the reader state was changed
        lngResult = SCardGetStatusChangeA(lngContext, 0.3, udtReaderStates(0), _
                                            lngReaderstatelen)
        DoEvents
        DoEvents
        DoEvents
        If (lngResult = 0) Then ' a reader presents
            'ReportCardStatus udtReaderStates(0).dwEventState
            'lblATRstring.Caption = ""
            Tester.Text1.Text = udtReaderStates(0).szReader
            Call AlcorCardTest
            
            'This reader was complete the test, wait for reader unplug
          '  Do
          '      DoEvents
          '      DoEvents
          '      DoEvents
          '      lngReaderstatelen = 1
          '      udtReaderStates(0).dwCurrentState = SCARD_STATE_UNAWARE
          '      ' check if the reader state was changed
          '      lngResult = SCardGetStatusChangeA(lngContext, 0, udtReaderStates(0), _
                                                lngReaderstatelen)
          '      If (lngResult <> 0) Then
          '          ShowResult ("CLEAR")
          '          txtmsg.Text = ""
          '          fnScsi2usb2K_KillEXE
          '          Exit Do
          '      End If
                    
          '  Loop While bStop = False
                
      '  Else ' no reader found
            
'            Call EvaluateResult("SCardGetStatusChangeA", lngResult)
         '   Text1.Text = "No Reader"
         '   Text2.Text = ""
           ' lblATRstring.Caption = ""
'            End
    
        End If
        DoEvents
        DoEvents
        DoEvents
   ' Loop While bStop = False
    
End Sub

Public Sub AlcorCardTest()
On Error Resume Next
    If (AlcorFindTheCard() = False) Then
        SmartCardTestResult = 2
        Exit Sub
    End If
    
    If (ProcessScriptTest() = False) Then
         SmartCardTestResult = 3
    Else
         SmartCardTestResult = 1
    End If
    
    
End Sub

Public Function AlcorFindTheCard() As Boolean
On Error Resume Next
    Dim lngReaderstatelen As Long
    Dim lngResult As Long
    Dim i As Long
    Dim strTmp As String
    
    AlcorFindTheCard = False
    
    Do
        lngReaderstatelen = 1
        udtReaderStates(0).dwCurrentState = SCARD_STATE_PRESENT 'SCARD_STATE_UNAWARE
        ' check if the reader state was changed
        lngResult = SCardGetStatusChangeA(lngContext, 0.2, udtReaderStates(0), _
                                                lngReaderstatelen)
        DoEvents
        DoEvents
        DoEvents
        'ReportCardStatus udtReaderStates(0).dwEventState
        If (lngResult <> 0) Then ' a reader not present,exit the sub
            OutMsg "Error! Can't find reader"
            Exit Do
        End If
        
        'If card is unrecognized, exit the sub
        If (udtReaderStates(0).dwEventState And SCARD_STATE_EMPTY) Then
'            OutMsg "No Card"
'            Exit Do
        ElseIf (udtReaderStates(0).dwEventState And SCARD_STATE_MUTE) Then
           OutMsg "Error! Unrecognized card"
           Exit Do
        
        Else
            
            For i = 0 To udtReaderStates(0).cbAtr - 1 'print the ATR string
                strTmp = strTmp & (Byte2Char(udtReaderStates(0).rgbAtr(i)) & " ")
            Next i
            OutMsg strTmp
            
            AlcorFindTheCard = True
            Exit Do
        End If
'        ReportCardStatus udtReaderStates(0).dwEventState
'        MsgBox udtReaderStates(0).rgbAtr
        
        DoEvents
        DoEvents
        DoEvents
        
        '    lblATRstring.Caption = ""

        
    Loop While bStop = False

End Function
Public Sub ShowResult(strResult As String)
    If ArchTest = True Then
    
        If (strResult = "PASS") Then
            Call LabelMenu(6, 1, 1)
            Tester.Label2.Caption = "PASS"
            Tester.Label2.ForeColor = vbGreen
            TestResultSmartCard = "PASS"
        ElseIf (strResult = "FAIL") Then
        Call LabelMenu(6, 0, 1)
            Tester.Label2.Caption = "FAIL"
            Tester.Label2.ForeColor = vbRed
             TestResultSmartCard = "FAIL"
        ElseIf (strResult = "CLEAR") Then
        Call LabelMenu(6, 0, 1)
            Tester.Label2.Caption = ""
            Tester.Label2.ForeColor = vbGreen
            TestResultSmartCard = "FAIL"
        End If
        
    Else

        If (strResult = "PASS") Then
            Call LabelMenu(5, 1, 1)
            Tester.Label2.Caption = "PASS"
            Tester.Label2.ForeColor = vbGreen
            TestResultadd = "PASS"
            ArchTest = True
        ElseIf (strResult = "FAIL") Then
        Call LabelMenu(5, 0, 1)
            Tester.Label2.Caption = "FAIL"
            Tester.Label2.ForeColor = vbRed
            TestResultadd = "FAIL"
            ArchTest = False
        ElseIf (strResult = "CLEAR") Then
        Call LabelMenu(5, 0, 1)
            Tester.Label2.Caption = ""
            Tester.Label2.ForeColor = vbGreen
             TestResultadd = "FAIL"
            ArchTest = False
        End If
    End If
    
End Sub


Public Function ProcessScriptTest() As Boolean

On Error Resume Next
    Dim strOneLine As String

    ProcessScriptTest = True
    Open strAppPath & "AlcorEMV.ini" For Input As #1
    lngLineNum = 0
    
    
    Do
        If (GetOneValidLine(strOneLine) = False) Then
            
            Exit Do
        End If
        
        Select Case (strOneLine)
            Case "[Card Protocol]"
                If (DoCardProtocol() = False) Then
                    ProcessScriptTest = False
                    Exit Do
                End If
               
            
            Case "[ATR]"
                If (CheckCardATR() = False) Then
                    OutMsg "Fail to check ATR"
                    ProcessScriptTest = False
                    Exit Do
                End If
                OutMsg "Check ATR PASS!"
            
            Case "[XFR]"
                If (DoCardXfrTest() = False) Then
                    OutMsg "Fail to Xfr Test"
                    ProcessScriptTest = False
                    Exit Do
                End If
                OutMsg "Check card command PASS"

            
            Case "[ChangeCard]"
                Tester.Label2.Caption = "Change card"
                Tester.Label2.ForeColor = vbBlue
                If (WaitCardChange() = False) Then
                    OutMsg "Error occurs when waiting card out"
                    ProcessScriptTest = False
                    Exit Do
                End If

                If (AlcorFindTheCard() = False) Then
                    OutMsg "Error occurs when waiting card out"
                    ProcessScriptTest = False
                    Exit Do
                End If
                
            Case Else
                OutMsg "Illigal line in AlcorEMV.ini, Line " & lngLineNum
                ProcessScriptTest = False
                Exit Do
        End Select
        
    Loop While (Not EOF(1))
    Close #1
    
End Function

Public Function GetOneValidLine(ByRef strOneLine As String) As Boolean
Dim Str As String
    While (Not EOF(1))
        Line Input #1, strOneLine
        lngLineNum = lngLineNum + 1
        strOneLine = Trim(strOneLine)
'        str = Mid(strOneLine, 1, 1)
        If (Mid(strOneLine, 1, 1) <> ";" And strOneLine <> "") Then
            strOneLine = strOneLine
            GetOneValidLine = True
            Exit Function
        End If
    Wend
    
    GetOneValidLine = False

End Function

Public Sub OutMsg(strMsg As String)
    Tester.txtmsg.Text = Tester.txtmsg.Text & strMsg & vbCrLf
    
End Sub


Public Function DoCardProtocol() As Boolean
    Dim strOneLine As String
    
    DoCardProtocol = False
    If (GetOneValidLine(strOneLine) = False) Then
        OutMsg "Can't get protocol setting in AlcorEMV.ini"
        Exit Function
    End If
    lngPreProtocol = 0
    Select Case (strOneLine)
        Case "Auto"
            lngPreProtocol = SCARD_PROTOCOL_T0 Or SCARD_PROTOCOL_T1
            DoCardProtocol = True
        Case "T0"
            lngPreProtocol = SCARD_PROTOCOL_T0
            DoCardProtocol = True
        Case "T1"
            lngPreProtocol = SCARD_PROTOCOL_T1
            DoCardProtocol = True
        Case Else
            OutMsg "Unknown protocol setting :" & strOneLine
            Exit Function
    End Select
    Call OutMsg("Card Protocol : " & strOneLine)
End Function

Public Function DoCardXfrTest() As Boolean
On Error Resume Next
    Dim i As Long
    Dim strCommand As String
    Dim lngCommandLen As Long
    Dim aCommand As ARRAY_TYPE
    Dim aResponse As ARRAY_TYPE
    Dim lngReplyLen As Long
    Dim lngReplyBufferLen As Long
    Dim lngResult As Long
    Dim pioSendPci As SCARD_IO_REQUEST
    Dim strReplyData As String
    Dim strOneLine As String
    Dim lngTestNo As Long
    Dim lNumOfTest As Long

    lngTestNo = 0
    DoCardXfrTest = False
    
    'Get the number of test to do
    If (GetOneValidLine(strOneLine) = False) Then
         OutMsg "Can't get protocol setting in AlcorEMV.ini"
    End If
    strOneLine = Trim(strOneLine)
    If (Mid(strOneLine, 1, 2) = "0x" Or Mid(strOneLine, 1, 2) = "0X") Then
        lNumOfTest = Val("&H" & Mid(strOneLine, 3, Len(strOneLine) - 2))
    Else
        lNumOfTest = strOneLine
    End If
    
    
    '=========================
    'Connect to the smartcard
    '=========================
    lngResult = SCardConnectA(lngContext, strReaderName, SCARD_SHARE_SHARED, lngPreProtocol, lngCard, lngActiveProtocol)
    
    If (lngResult <> 0) Then
        OutMsg "Can't connect to the card"
        Exit Function
    End If
'    If (VerifyAlcorReader() = False) Then
'        MsgBox "Not Alcor's reader"
'        Exit Function
'    End If
    
    lngResult = SCardBeginTransaction(lngCard)
    
    '=========================
    'Get command from file
    '=========================
    While (lngTestNo <> lNumOfTest)
        If (GetOneValidLine(strOneLine) = False) Then
            OutMsg "Invalid line in AlcorEMV.ini"
            Exit Function
        End If
        lngTestNo = lngTestNo + 1
        OutMsg "Command Test Case : " & lngTestNo
        strCommand = strOneLine
        lngCommandLen = (Len(strCommand) + 1) / 3

        For i = 1 To lngCommandLen
            aCommand.byteData(i - 1) = Val("&H" & Mid(strCommand, (i - 1) * 3 + 1, 2))
        Next i
        lngReplyLen = 258
        pioSendPci.dwProtocol = lngActiveProtocol
        pioSendPci.dbPciLength = Len(pioSendPci)
    
        lngResult = SCardTransmit(lngCard, pioSendPci, aCommand.byteData(0), lngCommandLen, pioSendPci, aResponse.byteData(0), lngReplyLen)
    
        DoEvents
        DoEvents
        DoEvents
        'txtCardResponse.Text = ""
        If (lngResult <> 0) Then
            OutMsg "Error! Fail to send smartcard data !"
            OutMsg "Cmd : " & strOneLine
            lngResult = SCardEndTransaction(lngCard, SCARD_LEAVE_CARD)
            Exit Function
        End If
        strReplyData = ""
        For i = 1 To lngReplyLen
            strReplyData = strReplyData & Byte2Char(aResponse.byteData(i - 1)) & " "
        Next i
        strReplyData = Trim(strReplyData)
        
        '=========================
        'Get the expected data from file
        '=========================
        If (GetOneValidLine(strOneLine) = False) Then
            OutMsg "Can't get protocol setting in AlcorEMV.ini"
            GoTo ExitFun
        End If
        
         '=========================
        'compare the result
        '=========================
        If (strOneLine <> strReplyData) Then
            OutMsg "Error! Card response compare error!"
            OutMsg "Expect : " & strOneLine
            OutMsg "Actual : " & strReplyData
            GoTo ExitFun
        End If
        DoEvents
        DoEvents
        DoEvents
    Wend
    
    DoCardXfrTest = True
    
ExitFun:
    lngResult = SCardEndTransaction(lngCard, SCARD_LEAVE_CARD)
    lngResult = SCardDisconnect(lngCard, SCARD_LEAVE_CARD)

End Function


Public Function ReportCardStatus(lngStatus As Long) As Long
    Dim strCardStatus As String
    
    strCardStatus = ""
    If (lngStatus And SCARD_STATE_IGNORE) Then
        strCardStatus = strCardStatus & "SCARD_STATE_IGNORE" & vbCrLf
    End If
    If (lngStatus And SCARD_STATE_CHANGED) Then
        strCardStatus = strCardStatus & "SCARD_STATE_CHANGED" & vbCrLf
    End If
    If (lngStatus And SCARD_STATE_UNKNOWN) Then
        strCardStatus = strCardStatus & "SCARD_STATE_UNKNOWN" & vbCrLf
    End If
    If (lngStatus And SCARD_STATE_UNAVAILABLE) Then
        strCardStatus = strCardStatus & "SCARD_STATE_UNAVAILABLE" & vbCrLf
    End If
    If (lngStatus And SCARD_STATE_EMPTY) Then
        strCardStatus = strCardStatus & "SCARD_STATE_EMPTY" & vbCrLf
    End If
    If (lngStatus And SCARD_STATE_PRESENT) Then
        strCardStatus = strCardStatus & "SCARD_STATE_PRESENT" & vbCrLf
    End If
    If (lngStatus And SCARD_STATE_ATRMATCH) Then
        strCardStatus = strCardStatus & "SCARD_STATE_ATRMATCH" & vbCrLf
    End If
    
    If (lngStatus And SCARD_STATE_EXCLUSIVE) Then
        strCardStatus = strCardStatus & "SCARD_STATE_EXCLUSIVE" & vbCrLf
    End If
    If (lngStatus And SCARD_STATE_INUSE) Then
        strCardStatus = strCardStatus & "SCARD_STATE_INUSE" & vbCrLf
    End If
    If (lngStatus And SCARD_STATE_MUTE) Then
        strCardStatus = strCardStatus & "SCARD_STATE_MUTE" & vbCrLf
    End If
    
    
    MsgBox strCardStatus
    
    

End Function


Public Function CheckCardATR() As Boolean
On Error Resume Next
    Dim i As Long
    Dim strTmp As String
    Dim strOneLine As String
    Dim lngReaderstatelen
    Dim lngResult As Long
    
    '=========================
    'Get ATR from smartcard service
    '=========================

    CheckCardATR = False
    lngReaderstatelen = 1
    udtReaderStates(0).dwCurrentState = SCARD_STATE_UNAWARE
    ' check if the reader state was changed
    lngResult = SCardGetStatusChangeA(lngContext, 100, udtReaderStates(0), _
                                                lngReaderstatelen)
    DoEvents
    DoEvents
    DoEvents
    'ReportCardStatus udtReaderStates(0).dwEventState
    If (lngResult <> 0 Or udtReaderStates(0).cbAtr = 0) Then ' a reader not present,exit the sub
        OutMsg "Fail to call SCardGetStatusChangeA in CheckCardATR()"
        Exit Function
    End If

    For i = 0 To udtReaderStates(0).cbAtr - 1 'print the ATR string
        strTmp = strTmp & (Byte2Char(udtReaderStates(0).rgbAtr(i)) & " ")
    Next i
    strTmp = Trim(strTmp)
    
    '=========================
    'Get ATR from file
    '=========================
    If (GetOneValidLine(strOneLine) = False) Then
        OutMsg "Can't get protocol setting in AlcorEMV.ini"
        Exit Function
    End If
   
    '=========================
    'Compare ATR
    '=========================
    If (strOneLine <> strTmp) Then
        OutMsg "The ATR not match!"
        OutMsg "Expect : " & strOneLine
        OutMsg "Actual : " & strTmp
        Exit Function
    End If

    CheckCardATR = True
        
End Function

Public Function VerifyAlcorReader() As Boolean
    
    Dim lngResult As Long
    Dim OutBuffer(1) As Byte
    Dim InBuffer(1) As Byte
    
    VerifyAlcorReader = False
    lngActiveProtocol = 0
    OutBuffer(0) = ASYNCHRONOUS_CARD_MODE
            
    lngResult = Alcor_SwitchCardMode(lngCard, 0, OutBuffer(0))
    If (lngResult <> 0) Then
        Exit Function
    End If
    
    VerifyAlcorReader = True

End Function


Private Sub EvaluateResult(strFnName As String, lngResult As Long)
On Error Resume Next
    ' output string
    Dim strOut As String
    ' string for text output of error value
    Dim strErrorValue As String
    ' leading zero padding to low hex values
    Dim strLeader As String
    Dim strMsg As String
    Dim lngOverallResult As Long
    Dim lngHexResultLen As Long
    strLeader = "00000000"
    
    ' print function name and pass/fail result
    strOut = strFnName
    
    If SCARD_S_SUCCESS = lngResult Then
        ' concatenate PASS tagline to output
 '       OutputLine strOut + strPass
    Else
        ' mark overall result flag
        lngOverallResult = SCARD_F_UNKNOWN_ERROR
        ' concatenate FAIL tagline to output
        strMsg = strOut + strFail
        ' convert result decimal number to hex string
        strErrorValue = Hex(lngResult)
        ' pad with leading zeros if necessary
        lngHexResultLen = Len(strErrorValue)
        strErrorValue = strTab + "Error: 0x" _
                        + Left(strLeader, 8 - lngHexResultLen) + strErrorValue
        strMsg = strMsg & vbCrLf & strErrorValue
        ' give message associated with error code
        strOut = ApiErrorMessage(lngResult)
        If 0 <> Len(strOut) Then
            strMsg = strMsg & vbCrLf & strTab + strOut
        End If
        MsgBox strMsg
    End If
    
End Sub ' EvaluateResult
Public Function WaitCardChange() As Boolean
On Error Resume Next
    Dim lngReaderstatelen As Long
    Dim lngResult As Long
    Dim i As Long
    Dim strTmp As String
    
    WaitCardChange = False
    
    Do
        lngReaderstatelen = 1
        udtReaderStates(0).dwCurrentState = SCARD_STATE_PRESENT 'SCARD_STATE_UNAWARE
        ' check if the reader state was changed
        lngResult = SCardGetStatusChangeA(lngContext, 0, udtReaderStates(0), _
                                                lngReaderstatelen)
        DoEvents
        DoEvents
        DoEvents
        'ReportCardStatus udtReaderStates(0).dwEventState
        If (lngResult <> 0) Then ' a reader not present,exit the sub
            OutMsg "Error! Can't find reader"
            Exit Do
        End If
        
        'If card is unrecognized, exit the sub
        If (udtReaderStates(0).dwEventState And SCARD_STATE_EMPTY) Then
            OutMsg "Find Card Out"
            
            WaitCardChange = True
            Exit Do
        End If
'        ReportCardStatus udtReaderStates(0).dwEventState
'        MsgBox udtReaderStates(0).rgbAtr
        
        DoEvents
        DoEvents
        DoEvents
        
        '    lblATRstring.Caption = ""

        
    Loop While bStop = False

End Function

