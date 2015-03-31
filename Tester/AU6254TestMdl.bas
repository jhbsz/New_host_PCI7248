Attribute VB_Name = "AU6254TestMdl"
Option Explicit
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
 
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Const SYNCHRONIZE = &H100000
Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const PROCESS_TERMINATE = &H1

Public Const SW_SHOW = 5

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Const WM_CLOSE = &H10
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type


Const SWP_NOMOVE = &H2 '不更動目前視窗位置
Const SWP_NOSIZE = &H1 '不更動目前視窗大小
Const HWND_TOPMOST = -1 '設定為最上層
Const HWND_NOTOPMOST = -2 '取消最上層設定
Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4
'ExitWindowsEx EWX_FORCE Or EWX_SHUTDOWN, 0

Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
'Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
'Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
'Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
'Const PROCESS_TERMINATE = (&H1)
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
   szExeFile            As String * MAX_PATH
End Type

Public AU6254CMediaFailCounter As Byte
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long


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

              CardResult = DO_WritePort(card, Channel_P1B, &HFF)
              Call MsecDelay(2)
              CardResult = DO_WritePort(card, Channel_P1B, &HF0)
                 
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

              CardResult = DO_WritePort(card, Channel_P1B, &HFF)
              Call MsecDelay(1)
              CardResult = DO_WritePort(card, Channel_P1B, &HF0)
                 
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

