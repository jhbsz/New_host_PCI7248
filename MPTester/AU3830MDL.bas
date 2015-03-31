Attribute VB_Name = "AU3830MDL"
Option Explicit

Public Const WM_USER = &H400
Public Const WM_CAM_MP_START = WM_USER + &H100
Public Const WM_CAM_MP_PASS = WM_USER + &H400
Public Const WM_CAM_MP_UNKNOW_FAIL = WM_USER + &H410
Public Const WM_CAM_MP_GPIO_FAIL = WM_USER + &H420
Public Const WM_CAM_MP_READY = WM_USER + &H800
 
Public Sub LoadMP_Click_AU3830()
    Dim TimePass
    Dim rt2
' find window
 

 
 
 
winHwnd = FindWindow(vbNullString, "VideoCap")
 
' run program
If winHwnd = 0 Then

Call ShellExecute(MPTester.hwnd, "open", App.Path & "\CamTest\" & ChipName & "\VideoCap.exe", "", "", SW_SHOW)
'Call ShellExecute(0, "open", App.Path & "\AlcorMP_698x_PD\AlcorMP.exe", "", "", SH_SHOW)
 
'Call ShellExecute(Me.hwnd, "open", App.Path & "\AlcorMP.exe", "", "", SH_SHOW)
 
End If

SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags
 End Sub

 
 Public Sub StartRWTest_Click_AU3830()
 Dim rt2
winHwnd = FindWindow(vbNullString, "VideoCap")
'debug.print "WindHandle="; winHwnd
rt2 = PostMessage(winHwnd, WM_CAM_MP_START, 0&, 0&)
 End Sub



 Public Sub AU3830A53ACF20TestSub()
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
 
   MPTester.TestResultLab = ""
'===============================================================
' Fail location initial
'===============================================================
 
 'AU7510 do not have filter driver
 
 
     
                     

AlcorMPMessage = 0
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
            Shell App.Path & "\CamTest\" & ChipName & "\ALInstFtr -i allow 3830"
            Call MsecDelay(1#)
           ' FileCopy App.Path & "\AlcorMP_698x_PD\RAM\" & chipname & "\RAM.Bin", App.Path & "\AlcorMP_698x_PD\RAM.Bin"
           ' FileCopy App.Path & "\AlcorMP_698x_PD\INI\" & chipname & "\AlcorMP.ini", App.Path & "\AlcorMP_698x_PD\AlcorMP.ini"
            NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName
 

 
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail


 '====================================
 '  Fix Card
 '====================================
 
 If NewChipFlag = 1 Or FindWindow(vbNullString, "VideoCap") = 0 Then
    
 '==============================================================
' when begin  scan + MP
'===============================================================
  
     
   
       '  power on
     '  cardresult = DO_WritePort(card, Channel_P1A, &HFF)
     '  Call PowerSet(3)   ' close power to disable chip
     '  Call MsecDelay(0.5)  ' power for load MPDriver
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
             
             '(2) initial fail :usb fail stage
      
             
         
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        Call StartRWTest_Click_AU3830
         
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
           '     Else
                 
           '       TestResult = "Bin3"
           '       MPTester.TestResultLab = "Bin3:LED FAIL "
              
           '    End If
               
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:Undefine Fail"
          
             ContFail = ContFail + 1
        
               
        End Select
         
        
         
                            
End Sub

