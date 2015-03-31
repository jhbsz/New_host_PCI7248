Attribute VB_Name = "AU6254TestMdl"

Public Function GetProcessId(Process As String) As Long
    Dim hSnapShot As Long, pe32 As PROCESSENTRY32
    hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, ByVal 0)
    pe32.dwSize = Len(pe32)
    ProcessFirst hSnapShot, pe32
    Do
        If InStr(1, pe32.szExeFile, Process & vbNullChar, vbTextCompare) = 1 Then
            GetProcessId = pe32.th32ProcessID
            Exit Do
        End If
    Loop While ProcessNext(hSnapShot, pe32)
    CloseHandle hSnapShot
End Function


Public Sub AU6254CMedia()
On Error Resume Next
Dim pId As Long, pHnd As Long
Dim winHwnd As Long
Dim winHwnd2 As Long
Dim winHwnd3 As Long
Dim RetVal As Long

Dim result
Dim TimeInterval
Dim TimeInterval2
Dim temp

Dim OldTimer2
Dim OldTimer
Dim TmpString As String

    Dim ProcessId As Long, hProcess As Long
    'Obtain the process id
    ProcessId = GetProcessId("RM5006.exe")
    'Obtain process handle
    If ProcessId <> 0 Then
    hProcess = OpenProcess(PROCESS_TERMINATE, 0, ProcessId)
    'Terminate the process
    TerminateProcess hProcess, 0
    'Close the process handle
    CloseHandle hProcess
    End If
 '1.==================================  Get AU6254
   
        If PCI7248InitFinish = 0 Then
                          PCI7248Exist
            result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
                          
        End If
           ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)

              cardresult = DO_WritePort(card, Channel_P1B, &HFF)
              Call MsecDelay(2)
              cardresult = DO_WritePort(card, Channel_P1B, &HF0)
                 
         '      CardResult = DO_WritePort(card, Channel_P1A, &H7F)
        '      Call MsecDelay(1.5)
            
        '    CardResult = DO_WritePort(card, Channel_P1A, &H7A)
        '    Call MsecDelay(0.3)
         '    CardResult = DO_WritePort(card, Channel_P1A, &HFA)
             
              Call MsecDelay(3)
                OldTimer = Timer
               Do
               DoEvents
                   rv0 = AU6254_GetDevice(0, 1, "6254")
                   TimeInterval = Timer - OldTimer
                   
               Loop While rv0 = 0 And TimeInterval < 5
                
                Tester.Print "rv0 ="; rv0; "TimeInterval="; TimeInterval
                
                Call LabelMenu(0, rv0, 1)
            
 '2.==================================  GetCMdia
 
               
                          
                If rv0 = 1 Then
       
                    TmpString = GetDeviceName("0d8c")
                    If TmpString <> "" Then
                      rv1 = 1
                      CMediaDiappearCounter = 0
                    Else
                      rv1 = 2
                       CMediaDiappearCounter = CMediaDiappearCounter + 1
                       
                        If CMediaDiappearCounter = 3 Then
                          Shell "shutdown -s -f -t 0"
                          End
                      End If
                    End If
                   
                Else
                   rv1 = 4
                End If
                 Call LabelMenu(1, rv1, rv0)
               
   '2.==================================  CMedia BurnIn test 30s
              
                If rv1 = 1 Then
                  
                 ' call CMedia Test program
                
                
                   Call ShellExecute(Tester.hwnd, "open", App.Path & "\RM5006.exe", "", "", SW_SHOW)
               
              
                
             '    pId = Shell(App.Path & "\RM5006.exe", vbNormalFocus)
                
                ' pHnd = OpenProcess(SYNCHRONIZE + PROCESS_QUERY_INFORMATION + PROCESS_TERMINATE, 0, pId)
                   
                    winHwnd = FindWindow(vbNullString, "RM5006 MP Test Program")
                    
                  
                   SetWindowPos winHwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags
                  
              
                OldTimer = Timer
               Do
               
               DoEvents
               temp = Timer
               Loop While temp - OldTimer < 3
               
               
               
                 Tester.Timer1.Interval = 500
                 Tester.Timer1.Enabled = True
                 
                   
              

                   OldTimer = Timer
                   CMediaPassCounter = 0
                   
                   Do
                   
                        DoEvents
                           winHwnd2 = FindWindow(vbNullString, "CSoundRec::Start")
                    
                    ' close CMedia test program
                  '  If winHwnd <> 0 Then
                      RetVal = PostMessage(winHwnd2, WM_CLOSE, 0&, 0&)
                      If RetVal = 0 Then
                        Shell "shutdown -s -f -t 0"
                     '   MsgBox "Error posting message."
                      End If
                      
                      
                        CMediaTestResult = 0
                        TimeInterval = Timer - OldTimer
                         OldTimer2 = Timer
                        Do
                         
                           DoEvents
                           'wait result from timer
                            TimeInterval2 = Timer - OldTimer2
                        Loop While CMediaTestResult = 0 And TimeInterval2 < 6
                         If CMediaTestResult = 1 And TimeInterval2 < 3 Then
                            CMediaTestResult = 0
                        End If
                        
                        If CMediaTestResult = 1 Then
                            CMediaPassCounter = CMediaPassCounter + 1
                        End If
                        Call MsecDelay(2)
                   Loop While CMediaTestResult = 1 And TimeInterval < 31   ' test cycle time
                   
                   Tester.Timer1.Enabled = False
                   
                   ' close CMedia test program
                    winHwnd3 = FindWindow(vbNullString, "RM5006 MP Test Program")
                    
                   
                     If winHwnd3 <> 0 Then
                      RetVal = PostMessage(winHwnd3, WM_CLOSE, 0&, 0&)
                      If RetVal = 0 Then
                        Shell "shutdown -s -f -t 0"
                        MsgBox "Error posting message."
                      End If
                 
                     End If
                  
                   ProcessId = GetProcessId("RM5006.exe")
                'Obtain process handle
                If ProcessId <> 0 Then
                hProcess = OpenProcess(PROCESS_TERMINATE, 0, ProcessId)
                'Terminate the process
                TerminateProcess hProcess, 0
                'Close the process handle
                CloseHandle hProcess
                End If
                
                    '========== get test result
                    Tester.Print CMediaPassCounter
                    If CMediaPassCounter > 5 Then    ' pass counter
                        rv2 = 1
                        Else
                        rv2 = 2
                    End If
                    
                Else
                   rv2 = 4
                End If
                
          Call LabelMenu(2, rv2, rv1)
                  
          
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
                ElseIf rv0 * rv1 * rv2 = 1 Then
                     TestResult = "PASS"
                     AU6254CMediaFailCounter = 0
                Else
                    TestResult = "Bin2"
                  
                End If
                
                If TestResult <> "PASS" Then
                
                   AU6254CMediaFailCounter = AU6254CMediaFailCounter + 1
                   
                   If AU6254CMediaFailCounter >= 5 Then
                       Shell "shutdown -s -f -t 0"
                   End If
                   
                End If
                
                
                
          '  CardResult = DO_WritePort(card, Channel_P1B, &HFF)
End Sub


Public Sub AU6254CMedia1()
On Error Resume Next
Dim pId As Long, pHnd As Long
Dim winHwnd As Long
Dim winHwnd2 As Long
Dim winHwnd3 As Long
Dim RetVal As Long


Dim TimeInterval
Dim TimeInterval2
Dim temp

Dim OldTimer2
Dim TmpString As String

    Dim ProcessId As Long, hProcess As Long
    'Obtain the process id
    ProcessId = GetProcessId("RM5006.exe")
    'Obtain process handle
    If ProcessId <> 0 Then
    hProcess = OpenProcess(PROCESS_TERMINATE, 0, ProcessId)
    'Terminate the process
    TerminateProcess hProcess, 0
    'Close the process handle
    CloseHandle hProcess
    End If
 '1.==================================  Get AU6254
   
        If PCI7248InitFinish = 0 Then
                          PCI7248Exist
                          result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
                          
        End If
           ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)

              cardresult = DO_WritePort(card, Channel_P1B, &HFF)
              Call MsecDelay(1)
              cardresult = DO_WritePort(card, Channel_P1B, &HF0)
                 
         '      CardResult = DO_WritePort(card, Channel_P1A, &H7F)
        '      Call MsecDelay(1.5)
            
        '    CardResult = DO_WritePort(card, Channel_P1A, &H7A)
        '    Call MsecDelay(0.3)
         '    CardResult = DO_WritePort(card, Channel_P1A, &HFA)
             
               Call MsecDelay(3)
                OldTimer = Timer
               Do
               DoEvents
                   rv0 = AU6254_GetDevice(0, 1, "6254")
                   TimeInterval = Timer - OldTimer
                   
               Loop While rv0 = 0 And TimeInterval < 5
                
                Tester.Print "rv0 ="; rv0; "TimeInterval="; TimeInterval
                
                Call LabelMenu(0, rv0, 1)
            
 '2.==================================  GetCMdia
 
               
                          
                If rv0 = 1 Then
       
                    TmpString = GetDeviceName("0d8c")
                    If TmpString <> "" Then
                      rv1 = 1
                      CMediaDiappearCounter = 0
                    Else
                      rv1 = 2
                       CMediaDiappearCounter = CMediaDiappearCounter + 1
                       
                        If CMediaDiappearCounter = 3 Then
                          Shell "shutdown -s -f -t 0"
                          End
                      End If
                    End If
                   
                Else
                   rv1 = 4
                End If
                 Call LabelMenu(1, rv1, rv0)
               
   '2.==================================  CMedia BurnIn test 30s
              
                If rv1 = 1 Then
                  
                 ' call CMedia Test program
                
                
                   Call ShellExecute(Tester.hwnd, "open", App.Path & "\RM5006.exe", "", "", SW_SHOW)
               
              
                
             '    pId = Shell(App.Path & "\RM5006.exe", vbNormalFocus)
                
                ' pHnd = OpenProcess(SYNCHRONIZE + PROCESS_QUERY_INFORMATION + PROCESS_TERMINATE, 0, pId)
                   
                   winHwnd = FindWindow(vbNullString, "RM5006 MP Test Program")
                    
                  SetForegroundWindow winHwnd
                '   SetWindowPos winHwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags
                  
              
                 OldTimer = Timer
                Do
               
                DoEvents
                temp = Timer
                Loop While temp - OldTimer < 1
               
               
               
                 Tester.Timer1.Interval = 500
                 Tester.Timer1.Enabled = True
                 
                   
              

                   OldTimer = Timer
                   CMediaPassCounter = 0
                   
                   Do
                   
                        DoEvents
                           winHwnd2 = FindWindow(vbNullString, "CSoundRec::Start")
                    
                    ' close CMedia test program
                  '  If winHwnd <> 0 Then
                      RetVal = PostMessage(winHwnd2, WM_CLOSE, 0&, 0&)
                      If RetVal = 0 Then
                        Shell "shutdown -s -f -t 0"
                     '   MsgBox "Error posting message."
                      End If
                      
                      
                        CMediaTestResult = 0
                        TimeInterval = Timer - OldTimer
                         OldTimer2 = Timer
                        Do
                         
                           DoEvents
                           'wait result from timer
                            TimeInterval2 = Timer - OldTimer2
                        Loop While CMediaTestResult = 0 And TimeInterval2 < 6
                         If CMediaTestResult = 1 And TimeInterval2 < 3 Then
                            CMediaTestResult = 0
                        End If
                        
                        If CMediaTestResult = 1 Then
                            CMediaPassCounter = CMediaPassCounter + 1
                        End If
                        Call MsecDelay(2)
                   Loop While CMediaTestResult = 1 And TimeInterval < 30   ' test cycle time
                   
                   Tester.Timer1.Enabled = False
                   
                   ' close CMedia test program
                    winHwnd3 = FindWindow(vbNullString, "RM5006 MP Test Program")
                    
                   
                     If winHwnd3 <> 0 Then
                      RetVal = PostMessage(winHwnd3, WM_CLOSE, 0&, 0&)
                      If RetVal = 0 Then
                        Shell "shutdown -s -f -t 0"
                        MsgBox "Error posting message."
                      End If
                 
                     End If
                  
                   ProcessId = GetProcessId("RM5006.exe")
                'Obtain process handle
                If ProcessId <> 0 Then
                hProcess = OpenProcess(PROCESS_TERMINATE, 0, ProcessId)
                'Terminate the process
                TerminateProcess hProcess, 0
                'Close the process handle
                CloseHandle hProcess
                End If
                
                    '========== get test result
                    Tester.Print CMediaPassCounter
                    If CMediaPassCounter > 5 Then    ' pass counter
                        rv2 = 1
                        Else
                        rv2 = 2
                    End If
                    
                Else
                   rv2 = 4
                End If
                
          Call LabelMenu(2, rv2, rv1)
                  
          
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
                ElseIf rv0 * rv1 * rv2 = 1 Then
                     TestResult = "PASS"
                     AU6254CMediaFailCounter = 0
                Else
                    TestResult = "Bin2"
                  
                End If
                
                If TestResult <> "PASS" Then
                
                   AU6254CMediaFailCounter = AU6254CMediaFailCounter + 1
                   
                   If AU6254CMediaFailCounter >= 5 Then
                       Shell "shutdown -s -f -t 0"
                   End If
                   
                End If
                
                
                
          '  CardResult = DO_WritePort(card, Channel_P1B, &HFF)
End Sub

