Attribute VB_Name = "AU6988MDL"

Public Sub AU6988D52ILF20TestSub()

Dim OldTimer
Dim PassTime
Dim rt2
Dim LightOn
Dim mMsg As MSG
Dim LedCount As Byte
 
    If PCI7248InitFinish = 0 Then
        PCI7248Exist
    End If
   
    MPTester.TestResultLab = ""
    '===============================================================
    ' Fail location initial
    '===============================================================
 
    If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
        FileCopy App.Path & "\AlcorMP_698x_PD\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
        Call MsecDelay(5)
    End If

    NewChipFlag = 0
    If OldChipName <> ChipName Then
        FileCopy App.Path & "\AlcorMP_698x_PD\ROM\" & ChipName & "\ROM.Hex", App.Path & "\AlcorMP_698x_PD\ROM.Hex"
        FileCopy App.Path & "\AlcorMP_698x_PD\RAM\" & ChipName & "\RAM.Bin", App.Path & "\AlcorMP_698x_PD\RAM.Bin"
        FileCopy App.Path & "\AlcorMP_698x_PD\INI\" & ChipName & "\AlcorMP.ini", App.Path & "\AlcorMP_698x_PD\AlcorMP.ini"
        NewChipFlag = 1 ' force MP
    End If
              
    OldChipName = ChipName
 
    '==============================================================
    ' when begin RW Test, must clear MP program
    '==============================================================
    
    '(1)
    winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
    If winHwnd <> 0 Then
        Do
            rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
            Call MsecDelay(0.5)
            winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
        Loop While winHwnd <> 0
    End If
    
    '(2)
    winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
    If winHwnd <> 0 Then
        Do
            rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
            Call MsecDelay(0.5)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
        Loop While winHwnd <> 0
    End If
    
    MPTester.Print "ContFail="; ContFail
    MPTester.Print "MPContFail="; MPContFail
 
    '====================================
    '  Fix Card
    '====================================
     
    If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag = True) Then
        
    '==============================================================
    ' when begin MP, must close RW program
    '==============================================================
        MPFlag = 1
 
        winHwnd = FindWindow(vbNullString, "UFD Test")
        If winHwnd <> 0 Then
            Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
            Loop While winHwnd <> 0
        End If
 
        '  power on
        cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        Call MsecDelay(0.5)  ' power for load MPDriver
        MPTester.Print "wait for MP Ready"
        Call LoadMP_Click_AU6988
 
        OldTimer = Timer
        AlcorMPMessage = 0
        
        Do
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
            PassTime = Timer - OldTimer
        Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 10 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3
        '====================================================
        If PassTime > 10 Then
            '(1)
            MPTester.Print "MP Ready Fail"
            winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
            If winHwnd <> 0 Then
                  Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                  Loop While winHwnd <> 0
            End If
            
           '(2)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
            If winHwnd <> 0 Then
                Do
                    rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                    Call MsecDelay(0.5)
                    winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                Loop While winHwnd <> 0
            End If
            
            MPTester.TestResultLab = "Bin3:MP Ready Fail"
            TestResult = "Bin3"
            MPTester.Print "MP Ready Fail"
     
            Exit Sub
        End If
        
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_MP_START Then
        
            cardresult = DO_WritePort(card, Channel_P1A, &H0)
            Call MsecDelay(6.5)
            MPTester.Print " MP Begin....."
            Call StartMP_Click_AU6988
            OldTimer = Timer
            AlcorMPMessage = 0
    
            Do
                If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                    AlcorMPMessage = mMsg.message
                    TranslateMessage mMsg
                    DispatchMessage mMsg
                End If
                
                PassTime = Timer - OldTimer
            Loop Until AlcorMPMessage = WM_FT_MP_PASS _
            Or AlcorMPMessage = WM_FT_MP_FAIL _
            Or AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL _
            Or PassTime > 60 _
            Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                
            MPTester.Print "MP work time="; PassTime
            MPTester.MPText.Text = Hex(AlcorMPMessage)
            
            '===============================================
            '  Handle MP work time out error
            '===============================================
                
            ' time out fail
            If PassTime > 60 Then
                MPContFail = MPContFail + 1
                TestResult = "Bin3"
                MPTester.TestResultLab = "Bin3:MP Time out Fail"
                MPTester.Print "MP Time out Fail"
                
                '(1)
                winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                If winHwnd <> 0 Then
                    Do
                        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                        Call MsecDelay(0.5)
                        winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    Loop While winHwnd <> 0
                End If
                
                '(2)
                winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                If winHwnd <> 0 Then
                    Do
                        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                        Call MsecDelay(0.5)
                        winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                    Loop While winHwnd <> 0
                End If
                Exit Sub
            End If
                
            ' MP fail
            If AlcorMPMessage = WM_FT_MP_FAIL Then
                MPContFail = MPContFail + 1
                TestResult = "Bin3"
                MPTester.TestResultLab = "Bin3:MP Function Fail"
                MPTester.Print "MP Function Fail"
                
                winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                If winHwnd <> 0 Then
                    Do
                        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                        Call MsecDelay(0.5)
                        winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    Loop While winHwnd <> 0
                End If
                
                winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                If winHwnd <> 0 Then
                    Do
                        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                        Call MsecDelay(0.5)
                        winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                    Loop While winHwnd <> 0
                End If
                Exit Sub
            End If
                
                
            'unknow fail
            If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                MPContFail = MPContFail + 1
                TestResult = "Bin2"
                MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                MPTester.Print "MP UNKNOW Fail"
                
                winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                If winHwnd <> 0 Then
                    Do
                        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                        Call MsecDelay(0.5)
                        winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    Loop While winHwnd <> 0
                End If
                 
                winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                If winHwnd <> 0 Then
                   Do
                       rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                       Call MsecDelay(0.5)
                       winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                   Loop While winHwnd <> 0
                End If
                Exit Sub
            End If
                 
                
            ' mp pass
            If AlcorMPMessage = WM_FT_MP_PASS Then
                MPTester.TestResultLab = "MP PASS"
                MPContFail = 0
                MPTester.Print "MP PASS"
            End If
        End If
    End If

    '=========================================
    '    Close MP program
    '=========================================
    winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
    If winHwnd <> 0 Then
        Do
            rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
            Call MsecDelay(0.5)
            winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
        Loop While winHwnd <> 0
    End If
    
    winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
    If winHwnd <> 0 Then
        Do
            rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
            Call MsecDelay(0.5)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
        Loop While winHwnd <> 0
    End If

                        
    '=========================================
    '    POWER on
    '=========================================
    If MPFlag = 1 Then
        cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        Call MsecDelay(0.5)  ' power of to unload MPDriver
        cardresult = DO_WritePort(card, Channel_P1A, &H0)
        Call MsecDelay(1.2)
        MPFlag = 0
    Else
        cardresult = DO_WritePort(card, Channel_P1A, &H0)
    End If
    
    Call LoadRWTest_Click_AU6988

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
        TestResult = "Bin3"
        MPTester.TestResultLab = "Bin3:RW Ready Fail"
       
        winHwnd = FindWindow(vbNullString, "UFD Test")
        If winHwnd <> 0 Then
            Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
            Loop While winHwnd <> 0
        End If
        Exit Sub
    End If

    OldTimer = Timer
    AlcorMPMessage = 0
    MPTester.Print "RW Tester begin test........"
    Call StartRWTest_Click_AU6988
        
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
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
                Do
                    rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                    Call MsecDelay(0.5)
                    winHwnd = FindWindow(vbNullString, "UFD Test")
                Loop While winHwnd <> 0
            End If
            Exit Sub
        End If
        
        Select Case AlcorMPMessage
        
            Case WM_FT_RW_UNKNOW_FAIL
                TestResult = "Bin2"
                MPTester.TestResultLab = "Bin2:UnKnow Fail"
                ContFail = ContFail + 1
            
            Case WM_FT_RW_SPEED_FAIL
                TestResult = "Bin3"
                MPTester.TestResultLab = "Bin3:SPEED Error "
                ContFail = ContFail + 1
                 
            Case WM_FT_RW_RW_FAIL
                TestResult = "Bin3"
                MPTester.TestResultLab = "Bin3:RW FAIL "
                ContFail = ContFail + 1
                 
            Case WM_FT_RW_ROM_FAIL
                TestResult = "Bin4"
                MPTester.TestResultLab = "Bin4:ROM FAIL "
                ContFail = ContFail + 1
                  
            Case WM_FT_RW_RAM_FAIL
                TestResult = "Bin5"
                MPTester.TestResultLab = "Bin5:RAM FAIL "
                ContFail = ContFail + 1
                
            Case WM_FT_RW_RW_PASS
                For LedCount = 1 To 20
                    Call MsecDelay(0.1)
                    cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
                    
                    If LightOn = 254 Then
                        Exit For
                    End If
                Next LedCount
                     
                MPTester.Print "light="; LightOn
                
                If LightOn = 254 Then
                    MPTester.TestResultLab = "PASS "
                    TestResult = "PASS"
                    ContFail = 0
                Else
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:LED FAIL "
                End If
 
            Case Else
                TestResult = "Bin2"
                MPTester.TestResultLab = "Bin2:Undefine Fail"
                ContFail = ContFail + 1
                
        End Select
                       
End Sub

Public Sub AU6988D52HLF22TestSub()

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
 
If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
    FileCopy App.Path & "\AlcorMP_698x_PD\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
    Call MsecDelay(5)
End If


NewChipFlag = 0
If OldChipName <> ChipName Then
            FileCopy App.Path & "\AlcorMP_698x_PD\ROM\" & ChipName & "\ROM.Hex", App.Path & "\AlcorMP_698x_PD\ROM.Hex"
            FileCopy App.Path & "\AlcorMP_698x_PD\RAM\" & ChipName & "\RAM.Bin", App.Path & "\AlcorMP_698x_PD\RAM.Bin"
            FileCopy App.Path & "\AlcorMP_698x_PD\INI\" & ChipName & "\AlcorMP.ini", App.Path & "\AlcorMP_698x_PD\AlcorMP.ini"
            NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName
 

'==============================================================
' when begin RW Test, must clear MP program
'===============================================================
'(1)
winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
  Loop While winHwnd <> 0
End If
'(2)
winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
  Loop While winHwnd <> 0
End If
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail
 
 
 '====================================
 '  Fix Card
 '====================================
 
 If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag = True) Then
    
 '==============================================================
' when begin MP, must close RW program
'===============================================================
   MPFlag = 1
 
    winHwnd = FindWindow(vbNullString, "UFD Test")
    If winHwnd <> 0 Then
      Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "UFD Test")
      Loop While winHwnd <> 0
    End If
 
       '  power on
       cardresult = DO_WritePort(card, Channel_P1A, &HFF)
       Call PowerSet(3)   ' close power to disable chip
       Call MsecDelay(0.5)  ' power for load MPDriver
       MPTester.Print "wait for MP Ready"
       Call LoadMP_Click_AU6988
 
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
        Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 10 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3
        '====================================================
        If PassTime > 10 Then
          '(1)
           MPTester.Print "MP Ready Fail"
            winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
              Loop While winHwnd <> 0
            End If
           '(2)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
            If winHwnd <> 0 Then
              Do
              rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
              Call MsecDelay(0.5)
              winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
              Loop While winHwnd <> 0
            End If
            
        
             MPTester.TestResultLab = "Bin3:MP Ready Fail"
             TestResult = "Bin3"
             MPTester.Print "MP Ready Fail"
     
              
            Exit Sub
        End If
        
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_MP_START Then
         
             
             cardresult = DO_WritePort(card, Channel_P1A, &HFD)  ' sel chip
             Call PowerSet(1900)   ' close power to disable chip
             Call MsecDelay(7.5)
               
             MPTester.Print " MP Begin....."
             
             Call StartMP_Click_AU6988
   
             
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
                
               ' time out fail
                If PassTime > 65 Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    '(1)
                    winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                      Loop While winHwnd <> 0
                    End If
                    
                    '(2)
                      winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    Exit Sub
                End If
                
                ' MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    
                    winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                      Loop While winHwnd <> 0
                    End If
                    
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    
                    Exit Sub
                End If
                
                
                 'unknow fail
                 If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                     MPContFail = MPContFail + 1
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                    MPTester.Print "MP UNKNOW Fail"
                    
                   winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                     Do
                     rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                     Call MsecDelay(0.5)
                     winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                     Loop While winHwnd <> 0
                     End If
                     
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                     If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                     End If
                     
                     
                     
                    Exit Sub
                End If
                 
                
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_PASS Then
                     MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPTester.Print "MP PASS"
                End If
        End If
   
End If
'=========================================
 '    Close MP program
 '=========================================
winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
If winHwnd <> 0 Then
  Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
  Loop While winHwnd <> 0
End If
    
 winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
    Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
    Loop While winHwnd <> 0
End If

                        
 '=========================================
 '    POWER on
 '=========================================
 If MPFlag = 1 Then
        Call PowerSet(3)
          cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        
         Call MsecDelay(0.5)  ' power of to unload MPDriver

           cardresult = DO_WritePort(card, Channel_P1A, &HF2)  'sel socket
          Call PowerSet(1900)
     
        
         Call MsecDelay(1.2)
        MPFlag = 0
 Else
          cardresult = DO_WritePort(card, Channel_P1A, &HF2)
         Call PowerSet(1900)
         
         Call MsecDelay(1.2)
End If
         Call LoadRWTest_Click_AU6988

        
        
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
        
       If PassTime > 5 Then
           TestResult = "Bin3"
           MPTester.TestResultLab = "Bin3:RW Ready Fail"
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
       
            Exit Sub
       End If
         
         
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        Call StartRWTest_Click_AU6988
        
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
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
            
              cardresult = DO_WritePort(card, Channel_P1A, &HF2)  ' power off
            
            Call PowerSet(1900)
        
            
       
            Exit Sub
        End If
        
        
     
               
        Select Case AlcorMPMessage
        
        Case WM_FT_RW_UNKNOW_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:UnKnow Fail"
          
             ContFail = ContFail + 1
        
        Case WM_FT_RW_SPEED_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:SPEED Error "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_RW_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:RW FAIL "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_ROM_FAIL
              TestResult = "Bin4"
              MPTester.TestResultLab = "Bin4:ROM FAIL "
              ContFail = ContFail + 1
              
        Case WM_FT_RW_RAM_FAIL
              TestResult = "Bin5"
              MPTester.TestResultLab = "Bin5:RAM FAIL "
               ContFail = ContFail + 1
        Case WM_FT_RW_RW_PASS
        
               
               For LedCount = 1 To 20
               Call MsecDelay(0.1)
               cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
                If LightOn = &HEF Or LightOn = &HCF Then
               
                 Exit For
               End If
               Next LedCount
                 
                  MPTester.Print "light="; LightOn
                 If LightOn = &HEF Or LightOn = &HCF Then
                    MPTester.TestResultLab = "PASS "
                    TestResult = "PASS"
                    ContFail = 0
                Else
                 
                  TestResult = "Bin3"
                  MPTester.TestResultLab = "Bin3:LED FAIL "
              
               End If
               
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:Undefine Fail"
          
             ContFail = ContFail + 1
        
               
        End Select
                               
       cardresult = DO_WritePort(card, Channel_P1A, &HF2)
        Call PowerSet(1900)
        
         
                            
End Sub


Public Sub AU6988D52HLF20TestSub()

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
 
If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
    FileCopy App.Path & "\AlcorMP_698x_PD\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
    Call MsecDelay(5)
End If


NewChipFlag = 0
If OldChipName <> ChipName Then
            FileCopy App.Path & "\AlcorMP_698x_PD\ROM\" & ChipName & "\ROM.Hex", App.Path & "\AlcorMP_698x_PD\ROM.Hex"
            FileCopy App.Path & "\AlcorMP_698x_PD\RAM\" & ChipName & "\RAM.Bin", App.Path & "\AlcorMP_698x_PD\RAM.Bin"
            FileCopy App.Path & "\AlcorMP_698x_PD\INI\" & ChipName & "\AlcorMP.ini", App.Path & "\AlcorMP_698x_PD\AlcorMP.ini"
            NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName
 

'==============================================================
' when begin RW Test, must clear MP program
'===============================================================
'(1)
winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
  Loop While winHwnd <> 0
End If
'(2)
winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
  Loop While winHwnd <> 0
End If
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail
 
 
 '====================================
 '  Fix Card
 '====================================
 
 If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag = True) Then
    
 '==============================================================
' when begin MP, must close RW program
'===============================================================
   MPFlag = 1
 
    winHwnd = FindWindow(vbNullString, "UFD Test")
    If winHwnd <> 0 Then
      Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "UFD Test")
      Loop While winHwnd <> 0
    End If
 
       '  power on
       cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        
       Call MsecDelay(0.5)  ' power for load MPDriver
       MPTester.Print "wait for MP Ready"
       Call LoadMP_Click_AU6988
 
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
        Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 10 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3
        '====================================================
        If PassTime > 10 Then
          '(1)
           MPTester.Print "MP Ready Fail"
            winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
              Loop While winHwnd <> 0
            End If
           '(2)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
            If winHwnd <> 0 Then
              Do
              rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
              Call MsecDelay(0.5)
              winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
              Loop While winHwnd <> 0
            End If
            
        
             MPTester.TestResultLab = "Bin3:MP Ready Fail"
             TestResult = "Bin3"
             MPTester.Print "MP Ready Fail"
     
              
            Exit Sub
        End If
        
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_MP_START Then
        
             cardresult = DO_WritePort(card, Channel_P1A, &HFB)
            
             Call MsecDelay(6.5)
               
             MPTester.Print " MP Begin....."
             
             Call StartMP_Click_AU6988
   
             
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
                Loop Until AlcorMPMessage = WM_FT_MP_PASS _
                Or AlcorMPMessage = WM_FT_MP_FAIL _
                Or AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL _
                Or PassTime > 60 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                
                MPTester.Print "MP work time="; PassTime
                 MPTester.MPText.Text = Hex(AlcorMPMessage)
                '================================================
                '  Handle MP work time out error
                '===============================================
                
               ' time out fail
                If PassTime > 60 Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    '(1)
                    winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                      Loop While winHwnd <> 0
                    End If
                    
                    '(2)
                      winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    Exit Sub
                End If
                
                ' MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    
                    winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                      Loop While winHwnd <> 0
                    End If
                    
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    
                    Exit Sub
                End If
                
                
                 'unknow fail
                 If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                     MPContFail = MPContFail + 1
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                    MPTester.Print "MP UNKNOW Fail"
                    
                   winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                     Do
                     rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                     Call MsecDelay(0.5)
                     winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                     Loop While winHwnd <> 0
                     End If
                     
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                     If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                     End If
                     
                     
                     
                    Exit Sub
                End If
                 
                
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_PASS Then
                     MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPTester.Print "MP PASS"
                End If
        End If
   
End If
'=========================================
 '    Close MP program
 '=========================================
winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
  Loop While winHwnd <> 0
End If
    
 winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
    Do
    rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
    Call MsecDelay(0.5)
    winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
    Loop While winHwnd <> 0
End If

                        
 '=========================================
 '    POWER on
 '=========================================
 If MPFlag = 1 Then
         cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        
         Call MsecDelay(0.5)  ' power of to unload MPDriver

         cardresult = DO_WritePort(card, Channel_P1A, &HFB)
        
         Call MsecDelay(1.2)
        MPFlag = 0
 Else
         cardresult = DO_WritePort(card, Channel_P1A, &HFB)
        
End If
         Call LoadRWTest_Click_AU6988

        
        
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
        
       If PassTime > 5 Then
           TestResult = "Bin3"
           MPTester.TestResultLab = "Bin3:RW Ready Fail"
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
       
            Exit Sub
       End If
         
         
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        Call StartRWTest_Click_AU6988
        
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
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
       
            Exit Sub
        End If
        
        
     
               
        Select Case AlcorMPMessage
        
        Case WM_FT_RW_UNKNOW_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:UnKnow Fail"
          
             ContFail = ContFail + 1
        
        Case WM_FT_RW_SPEED_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:SPEED Error "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_RW_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:RW FAIL "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_ROM_FAIL
              TestResult = "Bin4"
              MPTester.TestResultLab = "Bin4:ROM FAIL "
              ContFail = ContFail + 1
              
        Case WM_FT_RW_RAM_FAIL
              TestResult = "Bin5"
              MPTester.TestResultLab = "Bin5:RAM FAIL "
               ContFail = ContFail + 1
        Case WM_FT_RW_RW_PASS
        
               For LedCount = 1 To 20
               Call MsecDelay(0.1)
               cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
                If LightOn = 223 Then
               
                 Exit For
               End If
               Next LedCount
                 
                  MPTester.Print "light="; LightOn
                 If LightOn = 223 Then
                    MPTester.TestResultLab = "PASS "
                    TestResult = "PASS"
                    ContFail = 0
                Else
                 
                  TestResult = "Bin3"
                  MPTester.TestResultLab = "Bin3:LED FAIL "
              
               End If
               
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:Undefine Fail"
          
             ContFail = ContFail + 1
        
               
        End Select
                               
        '  cardresult = DO_WritePort(card, Channel_P1A, &HFF)
                            
End Sub

Public Sub AU6988D52HLF21TestSub()

 If PCI7248InitFinish = 0 Then
       PCI7248Exist
 End If
 Dim OldTimer
 Dim PassTime
 Dim rt2
 Dim LightOn
 Dim mMsg As MSG
 
   MPTester.TestResultLab = ""
'===============================================================
' Fail location initial
'===============================================================
 
If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
    FileCopy App.Path & "\AlcorMP_698x_PD\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
    Call MsecDelay(5)
End If


NewChipFlag = 0
If OldChipName <> ChipName Then
            FileCopy App.Path & "\AlcorMP_698x_PD\ROM\" & ChipName & "\ROM.Hex", App.Path & "\AlcorMP_698x_PD\ROM.Hex"
            FileCopy App.Path & "\AlcorMP_698x_PD\RAM\" & ChipName & "\RAM.Bin", App.Path & "\AlcorMP_698x_PD\RAM.Bin"
            FileCopy App.Path & "\AlcorMP_698x_PD\INI\" & ChipName & "\AlcorMP.ini", App.Path & "\AlcorMP_698x_PD\AlcorMP.ini"
            NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName
 

'==============================================================
' when begin RW Test, must clear MP program
'===============================================================
'(1)
winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
  Loop While winHwnd <> 0
End If
'(2)
winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
  Loop While winHwnd <> 0
End If
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail
 
 
 '====================================
 '  Fix Card
 '====================================
 
 If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag = True) Then
    
 '==============================================================
' when begin MP, must close RW program
'===============================================================
   MPFlag = 1
 
    winHwnd = FindWindow(vbNullString, "UFD Test")
    If winHwnd <> 0 Then
      Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "UFD Test")
      Loop While winHwnd <> 0
    End If
 
       '  power on
       cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        
       Call MsecDelay(0.5)  ' power for load MPDriver
       MPTester.Print "wait for MP Ready"
       Call LoadMP_Click_AU6988
 
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
        Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 10 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3
        '====================================================
        If PassTime > 10 Then
          '(1)
           MPTester.Print "MP Ready Fail"
            winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
              Loop While winHwnd <> 0
            End If
           '(2)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
            If winHwnd <> 0 Then
              Do
              rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
              Call MsecDelay(0.5)
              winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
              Loop While winHwnd <> 0
            End If
            
        
             MPTester.TestResultLab = "Bin3:MP Ready Fail"
             TestResult = "Bin3"
             MPTester.Print "MP Ready Fail"
     
              
            Exit Sub
        End If
        
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_MP_START Then
        
             cardresult = DO_WritePort(card, Channel_P1A, &HFB)
            
             Call MsecDelay(6.5)
               
             MPTester.Print " MP Begin....."
             
             Call StartMP_Click_AU6988
   
             
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
                Loop Until AlcorMPMessage = WM_FT_MP_PASS _
                Or AlcorMPMessage = WM_FT_MP_FAIL _
                Or AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL _
                Or PassTime > 55 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                
                MPTester.Print "MP work time="; PassTime
                 MPTester.MPText.Text = Hex(AlcorMPMessage)
                '================================================
                '  Handle MP work time out error
                '===============================================
                
               ' time out fail
                If PassTime > 55 Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    '(1)
                    winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                      Loop While winHwnd <> 0
                    End If
                    
                    '(2)
                      winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    Exit Sub
                End If
                
                ' MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    
                    winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                      Loop While winHwnd <> 0
                    End If
                    
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    
                    Exit Sub
                End If
                
                
                 'unknow fail
                 If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                     MPContFail = MPContFail + 1
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                    MPTester.Print "MP UNKNOW Fail"
                    
                   winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                     Do
                     rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                     Call MsecDelay(0.5)
                     winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                     Loop While winHwnd <> 0
                     End If
                     
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                     If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                     End If
                     
                     
                     
                    Exit Sub
                End If
                 
                
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_PASS Then
                     MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPTester.Print "MP PASS"
                End If
        End If
   
End If
'=========================================
 '    Close MP program
 '=========================================
winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
If winHwnd <> 0 Then
  Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
  Loop While winHwnd <> 0
End If
    
 winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
    Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
    Loop While winHwnd <> 0
End If

                        
 '=========================================
 '    POWER on
 '=========================================
 If MPFlag = 1 Then
         cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        
         Call MsecDelay(0.5)  ' power of to unload MPDriver

         cardresult = DO_WritePort(card, Channel_P1A, &HFB)
        
         Call MsecDelay(1.2)
        MPFlag = 0
 Else
         cardresult = DO_WritePort(card, Channel_P1A, &HFB)
         Call MsecDelay(1.2)
End If
         Call LoadRWTest_Click_AU6988

        
        
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
        
       If PassTime > 5 Then
           TestResult = "Bin3"
           MPTester.TestResultLab = "Bin3:RW Ready Fail"
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
       
            Exit Sub
       End If
         
         
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        Call StartRWTest_Click_AU6988
        
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
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
            
             cardresult = DO_WritePort(card, Channel_P1A, &HFF)   ' power off
        
            
       
            Exit Sub
        End If
        
        
     
               
        Select Case AlcorMPMessage
        
        Case WM_FT_RW_UNKNOW_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:UnKnow Fail"
          
             ContFail = ContFail + 1
        
        Case WM_FT_RW_SPEED_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:SPEED Error "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_RW_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:RW FAIL "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_ROM_FAIL
              TestResult = "Bin4"
              MPTester.TestResultLab = "Bin4:ROM FAIL "
              ContFail = ContFail + 1
              
        Case WM_FT_RW_RAM_FAIL
              TestResult = "Bin5"
              MPTester.TestResultLab = "Bin5:RAM FAIL "
               ContFail = ContFail + 1
        Case WM_FT_RW_RW_PASS
        
               
               cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
               MPTester.Print "light="; LightOn
              
               If LightOn = 223 Then
                    MPTester.TestResultLab = "PASS "
                    TestResult = "PASS"
                    ContFail = 0
               Else
                     Call MsecDelay(0.5)
                     cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
                     MPTester.Print "light="; LightOn
                     If LightOn = 223 Then
                        MPTester.TestResultLab = "PASS "
                        TestResult = "PASS"
                        ContFail = 0
                    Else
                        TestResult = "Bin3"
                        MPTester.TestResultLab = "Bin3:LED FAIL "
                    End If
               End If
               
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:Undefine Fail"
          
             ContFail = ContFail + 1
        
               
        End Select
                               
         cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        
         
                            
End Sub
Public Sub AU6988D52HLF23TestSub()
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
 
If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
    FileCopy App.Path & "\AlcorMP_698x_PD\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
    Call MsecDelay(5)
End If


NewChipFlag = 0
If OldChipName <> ChipName Then
            FileCopy App.Path & "\AlcorMP_698x_PD\ROM\" & ChipName & "\ROM.Hex", App.Path & "\AlcorMP_698x_PD\ROM.Hex"
            FileCopy App.Path & "\AlcorMP_698x_PD\RAM\" & ChipName & "\RAM.Bin", App.Path & "\AlcorMP_698x_PD\RAM.Bin"
            FileCopy App.Path & "\AlcorMP_698x_PD\INI\" & ChipName & "\AlcorMP.ini", App.Path & "\AlcorMP_698x_PD\AlcorMP.ini"
            NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName
 

'==============================================================
' when begin RW Test, must clear MP program
'===============================================================
'(1)
winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
  Loop While winHwnd <> 0
End If
'(2)
winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
  Loop While winHwnd <> 0
End If
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail
 
 
 '====================================
 '  Fix Card
 '====================================
 
 If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag = True) Then
    
 '==============================================================
' when begin MP, must close RW program
'===============================================================
   MPFlag = 1
 
    winHwnd = FindWindow(vbNullString, "UFD Test")
    If winHwnd <> 0 Then
      Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "UFD Test")
      Loop While winHwnd <> 0
    End If
 
       '  power on
       cardresult = DO_WritePort(card, Channel_P1A, &HFF)
       Call PowerSet(3)   ' close power to disable chip
       Call MsecDelay(0.5)  ' power for load MPDriver
       MPTester.Print "wait for MP Ready"
       Call LoadMP_Click_AU6988
 
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
        Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 10 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3
        '====================================================
        If PassTime > 10 Then
          '(1)
           MPTester.Print "MP Ready Fail"
            winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
              Loop While winHwnd <> 0
            End If
           '(2)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
            If winHwnd <> 0 Then
              Do
              rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
              Call MsecDelay(0.5)
              winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
              Loop While winHwnd <> 0
            End If
            
        
             MPTester.TestResultLab = "Bin3:MP Ready Fail"
             TestResult = "Bin3"
             MPTester.Print "MP Ready Fail"
     
              
            Exit Sub
        End If
        
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_MP_START Then
         
             
             cardresult = DO_WritePort(card, Channel_P1A, &HFD)  ' sel chip
             Call PowerSet(1900)   ' close power to disable chip
             Call MsecDelay(7.5)
               
             MPTester.Print " MP Begin....."
             
             Call StartMP_Click_AU6988
   
             
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
                
               ' time out fail
                If PassTime > 65 Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    '(1)
                    winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                      Loop While winHwnd <> 0
                    End If
                    
                    '(2)
                      winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    Exit Sub
                End If
                
                ' MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    
                    winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                      Loop While winHwnd <> 0
                    End If
                    
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    
                    Exit Sub
                End If
                
                
                 'unknow fail
                 If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                     MPContFail = MPContFail + 1
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                    MPTester.Print "MP UNKNOW Fail"
                    
                   winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                     Do
                     rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                     Call MsecDelay(0.5)
                     winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                     Loop While winHwnd <> 0
                     End If
                     
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                     If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                     End If
                     
                     
                     
                    Exit Sub
                End If
                 
                
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_PASS Then
                     MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPTester.Print "MP PASS"
                End If
        End If
   
End If
'=========================================
 '    Close MP program
 '=========================================
winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
If winHwnd <> 0 Then
  Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
  Loop While winHwnd <> 0
End If
    
 winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
    Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
    Loop While winHwnd <> 0
End If

  Dim pid As Long          ' unload driver
  Dim hProcess As Long
  Dim ExitEvent As Long
 
  pid = Shell(App.Path & "\AlcorMP_698x_PD\loaddrv.exe uninstall_058F6387")
  hProcess = OpenProcess(SYNCHRONIZE + PROCESS_QUERY_INFORMATION + PROCESS_TERMINATE, 0, pid)
  ExitEvent = WaitForSingleObject(hProcess, INFINITE)
  Call CloseHandle(hProcess)
 
 
 
                        
 '=========================================
 '    POWER on
 '=========================================
 
 If MPFlag = 1 Then
        Call PowerSet(3)
          cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        
         Call MsecDelay(0.5)  ' power of to unload MPDriver

           cardresult = DO_WritePort(card, Channel_P1A, &HF2)  'sel socket
          Call PowerSet(1900)
     
        
         Call MsecDelay(1.2)
        MPFlag = 0
 Else
          cardresult = DO_WritePort(card, Channel_P1A, &HF2)
         Call PowerSet(1900)
         
         Call MsecDelay(1.2)
End If
         Call LoadRWTest_Click_AU6988

        
        
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
        
       If PassTime > 5 Then
           TestResult = "Bin3"
           MPTester.TestResultLab = "Bin3:RW Ready Fail"
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
       
            Exit Sub
       End If
         
         
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        Call StartRWTest_Click_AU6988
        
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
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
            
              cardresult = DO_WritePort(card, Channel_P1A, &HF2)  ' power off
            
            Call PowerSet(1900)
        
            
       
            Exit Sub
        End If
        
        
     
               
        Select Case AlcorMPMessage
        
        Case WM_FT_RW_UNKNOW_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:UnKnow Fail"
          
             ContFail = ContFail + 1
        
        Case WM_FT_RW_SPEED_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:SPEED Error "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_RW_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:RW FAIL "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_ROM_FAIL
              TestResult = "Bin4"
              MPTester.TestResultLab = "Bin4:ROM FAIL "
              ContFail = ContFail + 1
              
        Case WM_FT_RW_RAM_FAIL
              TestResult = "Bin5"
              MPTester.TestResultLab = "Bin5:RAM FAIL "
               ContFail = ContFail + 1
        Case WM_FT_RW_RW_PASS
        
               
               For LedCount = 1 To 20
               Call MsecDelay(0.1)
               cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
                If LightOn = &HEF Or LightOn = &HCF Then
               
                 Exit For
               End If
               Next LedCount
                 
                  MPTester.Print "light="; LightOn
                 If LightOn = &HEF Or LightOn = &HCF Then
                    MPTester.TestResultLab = "PASS "
                    TestResult = "PASS"
                    ContFail = 0
                Else
                 
                  TestResult = "Bin3"
                  MPTester.TestResultLab = "Bin3:LED FAIL "
              
               End If
               
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:Undefine Fail"
          
             ContFail = ContFail + 1
        
               
        End Select
                               
       cardresult = DO_WritePort(card, Channel_P1A, &HF2)
        Call PowerSet(1900)
        
         
                            
End Sub

Public Sub AU6988D52HLF24TestSub()
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
 
If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
    FileCopy App.Path & "\AlcorMP_698x_PD\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
    Call MsecDelay(5)
End If


NewChipFlag = 0
If OldChipName <> ChipName Then
            FileCopy App.Path & "\AlcorMP_698x_PD\ROM\" & ChipName & "\ROM.Hex", App.Path & "\AlcorMP_698x_PD\ROM.Hex"
            FileCopy App.Path & "\AlcorMP_698x_PD\RAM\" & ChipName & "\RAM.Bin", App.Path & "\AlcorMP_698x_PD\RAM.Bin"
            FileCopy App.Path & "\AlcorMP_698x_PD\INI\" & ChipName & "\AlcorMP.ini", App.Path & "\AlcorMP_698x_PD\AlcorMP.ini"
            NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName
 

'==============================================================
' when begin RW Test, must clear MP program
'===============================================================
'(1)
winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
  Loop While winHwnd <> 0
End If
'(2)
winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
  Loop While winHwnd <> 0
End If
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail
 
 
 '====================================
 '  Fix Card
 '====================================
 
 If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag = True) Then
    
 '==============================================================
' when begin MP, must close RW program
'===============================================================
   MPFlag = 1
 
    winHwnd = FindWindow(vbNullString, "UFD Test")
    If winHwnd <> 0 Then
      Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "UFD Test")
      Loop While winHwnd <> 0
    End If
 
       '  power on
       cardresult = DO_WritePort(card, Channel_P1A, &HFF)
       Call PowerSet(3)   ' close power to disable chip
       Call MsecDelay(0.5)  ' power for load MPDriver
       MPTester.Print "wait for MP Ready"
       Call LoadMP_Click_AU6988
 
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
        Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 30 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3
        '====================================================
        If PassTime > 30 Then
          '(1)
           MPTester.Print "MP Ready Fail"
            winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
              Loop While winHwnd <> 0
            End If
           '(2)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
            If winHwnd <> 0 Then
              Do
              rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
              Call MsecDelay(0.5)
              winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
              Loop While winHwnd <> 0
            End If
            
        
             MPTester.TestResultLab = "Bin3:MP Ready Fail"
             TestResult = "Bin3"
             MPTester.Print "MP Ready Fail"
     
              
            Exit Sub
        End If
        
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_MP_START Then
         
             
             cardresult = DO_WritePort(card, Channel_P1A, &HFD)  ' sel chip
             Call PowerSet(500)   ' close power to disable chip
             Call MsecDelay(7.5)
               
             MPTester.Print " MP Begin....."
             
             Call StartMP_Click_AU6988
   
             
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
                
               ' time out fail
                If PassTime > 65 Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    '(1)
                    winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                      Loop While winHwnd <> 0
                    End If
                    
                    '(2)
                      winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    Exit Sub
                End If
                
                ' MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    
                    winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                      Loop While winHwnd <> 0
                    End If
                    
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    
                    Exit Sub
                End If
                
                
                 'unknow fail
                 If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                     MPContFail = MPContFail + 1
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                    MPTester.Print "MP UNKNOW Fail"
                    
                   winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                     Do
                     rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                     Call MsecDelay(0.5)
                     winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                     Loop While winHwnd <> 0
                     End If
                     
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                     If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                     End If
                     
                     
                     
                    Exit Sub
                End If
                 
                
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_PASS Then
                     MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPTester.Print "MP PASS"
                End If
        End If
   
End If
'=========================================
 '    Close MP program
 '=========================================
winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
If winHwnd <> 0 Then
  Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
  Loop While winHwnd <> 0
End If
    
 winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
    Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
    Loop While winHwnd <> 0
End If

  Dim pid As Long          ' unload driver
  Dim hProcess As Long
  Dim ExitEvent As Long
 
  pid = Shell(App.Path & "\AlcorMP_698x_PD\loaddrv.exe uninstall_058F6387")
  hProcess = OpenProcess(SYNCHRONIZE + PROCESS_QUERY_INFORMATION + PROCESS_TERMINATE, 0, pid)
  ExitEvent = WaitForSingleObject(hProcess, INFINITE)
  Call CloseHandle(hProcess)
 
 
 
                        
 '=========================================
 '    POWER on
 '=========================================
 
 If MPFlag = 1 Then
        Call PowerSet(3)
          cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        
         Call MsecDelay(0.5)  ' power of to unload MPDriver

           cardresult = DO_WritePort(card, Channel_P1A, &HF2)  'sel socket
          Call PowerSet(1500)
     
        
         Call MsecDelay(1.2)
        MPFlag = 0
 Else
          cardresult = DO_WritePort(card, Channel_P1A, &HF2)
         Call PowerSet(1500)
         
         Call MsecDelay(1.2)
End If
         Call LoadRWTest_Click_AU6988

        
        
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
        
       If PassTime > 5 Then
           TestResult = "Bin3"
           MPTester.TestResultLab = "Bin3:RW Ready Fail"
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
       
            Exit Sub
       End If
         
         
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        Call StartRWTest_Click_AU6988
        
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
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
            
              cardresult = DO_WritePort(card, Channel_P1A, &HF2)  ' power off
            
            Call PowerSet(1500)
        
            
       
            Exit Sub
        End If
        
        
     
               
        Select Case AlcorMPMessage
        
        Case WM_FT_RW_UNKNOW_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:UnKnow Fail"
          
             ContFail = ContFail + 1
        
        Case WM_FT_RW_SPEED_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:SPEED Error "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_RW_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:RW FAIL "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_ROM_FAIL
              TestResult = "Bin4"
              MPTester.TestResultLab = "Bin4:ROM FAIL "
              ContFail = ContFail + 1
              
        Case WM_FT_RW_RAM_FAIL
              TestResult = "Bin5"
              MPTester.TestResultLab = "Bin5:RAM FAIL "
               ContFail = ContFail + 1
        Case WM_FT_RW_RW_PASS
        
               
               For LedCount = 1 To 20
               Call MsecDelay(0.1)
               cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
                If LightOn = &HEF Or LightOn = &HCF Then
               
                 Exit For
               End If
               Next LedCount
                 
                  MPTester.Print "light="; LightOn
                 If LightOn = &HEF Or LightOn = &HCF Then
                    MPTester.TestResultLab = "PASS "
                    TestResult = "PASS"
                    ContFail = 0
                Else
                 
                  TestResult = "Bin3"
                  MPTester.TestResultLab = "Bin3:LED FAIL "
              
               End If
               
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:Undefine Fail"
          
             ContFail = ContFail + 1
        
               
        End Select
                               
       cardresult = DO_WritePort(card, Channel_P1A, &HF2)
        Call PowerSet(1500)
        
         
                            
End Sub

Public Sub AU6988D52HLF25TestSub()
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
Dim TimerCounter As Integer
Dim TmpString As String

   MPTester.TestResultLab = ""
'===============================================================
' Fail location initial
'===============================================================
 
If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
    FileCopy App.Path & "\AlcorMP_698x_PD\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
    Call MsecDelay(5)
End If


NewChipFlag = 0
If OldChipName <> ChipName Then
            FileCopy App.Path & "\AlcorMP_698x_PD\ROM\" & ChipName & "\ROM.Hex", App.Path & "\AlcorMP_698x_PD\ROM.Hex"
            FileCopy App.Path & "\AlcorMP_698x_PD\RAM\" & ChipName & "\RAM.Bin", App.Path & "\AlcorMP_698x_PD\RAM.Bin"
            FileCopy App.Path & "\AlcorMP_698x_PD\INI\" & ChipName & "\AlcorMP.ini", App.Path & "\AlcorMP_698x_PD\AlcorMP.ini"
            NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName
 

'==============================================================
' when begin RW Test, must clear MP program
'===============================================================
'(1)
winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
  Loop While winHwnd <> 0
End If
'(2)
winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
  Loop While winHwnd <> 0
End If
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail
 
 
 '====================================
 '  Fix Card
 '====================================
 
 If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag = True) Then
    
 '==============================================================
' when begin MP, must close RW program
'===============================================================
   MPFlag = 1
 
    winHwnd = FindWindow(vbNullString, "UFD Test")
    If winHwnd <> 0 Then
      Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "UFD Test")
      Loop While winHwnd <> 0
    End If
 
       '  power on
       cardresult = DO_WritePort(card, Channel_P1A, &HFF)
       Call PowerSet(3)   ' close power to disable chip
       Call MsecDelay(0.5)  ' power for load MPDriver
       MPTester.Print "wait for MP Ready"
       Call LoadMP_Click_AU6988
 
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
        Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 30 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3
        '====================================================
        If PassTime > 30 Then
          '(1)
           MPTester.Print "MP Ready Fail"
            winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
              Loop While winHwnd <> 0
            End If
           '(2)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
            If winHwnd <> 0 Then
              Do
              rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
              Call MsecDelay(0.5)
              winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
              Loop While winHwnd <> 0
            End If
            
        
             MPTester.TestResultLab = "Bin3:MP Ready Fail"
             TestResult = "Bin3"
             MPTester.Print "MP Ready Fail"
     
              
            Exit Sub
        End If
        
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_MP_START Then
         
             
             cardresult = DO_WritePort(card, Channel_P1A, &HFD)  ' sel chip
             Call PowerSet(500)   ' close power to disable chip
             
 
           '  Call MsecDelay(6.5)
            
             
            Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName("vid")
            Loop While TmpString = "" And TimerCounter < 150
             
             
             If TmpString = "" Then   ' can not find device after 15 s
             
               TestResult = "Bin2"
               MPTester.TestResultLab = "Bin2:MP UNKNOW Fail when enter MP"
               Exit Sub
             End If
             
             Call MsecDelay(2.5)
               
             MPTester.Print " MP Begin....."
             
             Call StartMP_Click_AU6988
   
             
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
                
               ' time out fail
                If PassTime > 65 Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    '(1)
                    winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                      Loop While winHwnd <> 0
                    End If
                    
                    '(2)
                      winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    Exit Sub
                End If
                
                ' MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    
                    winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                      Loop While winHwnd <> 0
                    End If
                    
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    
                    Exit Sub
                End If
                
                
                 'unknow fail
                 If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                     MPContFail = MPContFail + 1
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                    MPTester.Print "MP UNKNOW Fail"
                    
                   winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                     Do
                     rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                     Call MsecDelay(0.5)
                     winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                     Loop While winHwnd <> 0
                     End If
                     
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                     If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                     End If
                     
                     
                     
                    Exit Sub
                End If
                 
                
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_PASS Then
                     MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPTester.Print "MP PASS"
                End If
        End If
   
End If
'=========================================
 '    Close MP program
 '=========================================
winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
If winHwnd <> 0 Then
  Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
  Loop While winHwnd <> 0
End If
    
 winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
    Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
    Loop While winHwnd <> 0
End If

  Dim pid As Long          ' unload driver
  Dim hProcess As Long
  Dim ExitEvent As Long
 
  pid = Shell(App.Path & "\AlcorMP_698x_PD\loaddrv.exe uninstall_058F6387")
  hProcess = OpenProcess(SYNCHRONIZE + PROCESS_QUERY_INFORMATION + PROCESS_TERMINATE, 0, pid)
  ExitEvent = WaitForSingleObject(hProcess, INFINITE)
  Call CloseHandle(hProcess)
 
 
 
                        
 '=========================================
 '    POWER on
 '=========================================
 
 If MPFlag = 1 Then
        Call PowerSet(3)
          cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        
         Call MsecDelay(0.5)  ' power of to unload MPDriver

           cardresult = DO_WritePort(card, Channel_P1A, &HF2)  'sel socket
          Call PowerSet(1500)
     
        
         Call MsecDelay(1.2)
        MPFlag = 0
 Else
          cardresult = DO_WritePort(card, Channel_P1A, &HF2)
         Call PowerSet(1500)
         
         Call MsecDelay(1.2)
End If
         Call LoadRWTest_Click_AU6988

        
        
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
        
       If PassTime > 5 Then
           TestResult = "Bin3"
           MPTester.TestResultLab = "Bin3:RW Ready Fail"
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
       
            Exit Sub
       End If
         
         
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        Call StartRWTest_Click_AU6988
        
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
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
            
              cardresult = DO_WritePort(card, Channel_P1A, &HF2)  ' power off
            
            Call PowerSet(1500)
        
            
       
            Exit Sub
        End If
        
        
     
               
        Select Case AlcorMPMessage
        
        Case WM_FT_RW_UNKNOW_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:UnKnow Fail"
          
             ContFail = ContFail + 1
        
        Case WM_FT_RW_SPEED_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:SPEED Error "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_RW_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:RW FAIL "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_ROM_FAIL
              TestResult = "Bin4"
              MPTester.TestResultLab = "Bin4:ROM FAIL "
              ContFail = ContFail + 1
              
        Case WM_FT_RW_RAM_FAIL
              TestResult = "Bin5"
              MPTester.TestResultLab = "Bin5:RAM FAIL "
               ContFail = ContFail + 1
        Case WM_FT_RW_RW_PASS
        
               
               For LedCount = 1 To 20
               Call MsecDelay(0.1)
               cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
                If LightOn = &HEF Or LightOn = &HCF Then
               
                 Exit For
               End If
               Next LedCount
                 
                  MPTester.Print "light="; LightOn
                 If LightOn = &HEF Or LightOn = &HCF Then
                    MPTester.TestResultLab = "PASS "
                    TestResult = "PASS"
                    ContFail = 0
                Else
                 
                  TestResult = "Bin3"
                  MPTester.TestResultLab = "Bin3:LED FAIL "
              
               End If
               
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:Undefine Fail"
          
             ContFail = ContFail + 1
        
               
        End Select
                               
       cardresult = DO_WritePort(card, Channel_P1A, &HF2)
        Call PowerSet(1500)
        
         
                            
End Sub

Public Sub AU6988D52HLF2ITestSub()

'Support K9F1G
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
 'Dim ReMP_Flag As Byte
 
 
 
   MPTester.TestResultLab = ""
'===============================================================
' Fail loacatio initial
'===============================================================
 
If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
    FileCopy App.Path & "\AlcorMP_6988\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
    Call MsecDelay(5)
End If


NewChipFlag = 0
If OldChipName <> ChipName Then
    FileCopy App.Path & "\AlcorMP_6988\INI\" & ChipName & "\ROM.Hex", App.Path & "\AlcorMP_6988\ROM.Hex"
    FileCopy App.Path & "\AlcorMP_6988\INI\" & ChipName & "\RAM.Bin", App.Path & "\AlcorMP_6988\RAM.Bin"
    FileCopy App.Path & "\AlcorMP_6988\INI\" & ChipName & "\AlcorMP.ini", App.Path & "\AlcorMP_6988\AlcorMP.ini"
    FileCopy App.Path & "\AlcorMP_6988\INI\" & ChipName & "\PE.bin", App.Path & "\AlcorMP_6988\PE.bin"
    FileCopy App.Path & "\AlcorMP_6988\INI\" & ChipName & "\FT.ini", App.Path & "\FT.ini"
    FileCopy App.Path & "\AlcorMP_6988\INI\" & ChipName & "\FT.ini", App.Path & "\AlcorMP_6988\FT.ini"
    NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName
 

'==============================================================
' when begin RW Test, must clear MP porgram
'===============================================================


'(1)  /// for Auto mode
winHwnd = FindWindow(vbNullString, AU6988MPCaption1)
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, AU6988MPCaption1)
  Loop While winHwnd <> 0
End If

'(1)
winHwnd = FindWindow(vbNullString, AU6988MPCaption)
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, AU6988MPCaption)
  Loop While winHwnd <> 0
End If
'(2)
winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
  Loop While winHwnd <> 0
End If
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail
 
 
 '====================================
 '  Fix Card
 '====================================
' GoTo T1
 If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag = True) Then
 
 
   If MPTester.NoMP.Value = 1 Then
        
        If (NewChipFlag = 0) And (MPTester.Check1.Value = 0) Then  ' force condition
           GoTo RW_Test_Label
        End If
    End If
       
    
    If MPTester.ResetMPFailCounter.Value = 1 Then
       ContFail = 0
    End If
    
 '==============================================================
' when begin MP, must close RW porgram
'===============================================================
   MPFlag = 1
 
    winHwnd = FindWindow(vbNullString, "UFD Test")
    If winHwnd <> 0 Then
      Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "UFD Test")
      Loop While winHwnd <> 0
    End If
 
       '  power on
       cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        Call PowerSet(3)   ' close power to disable chip
       Call MsecDelay(2.5)  ' power for load MPDriver
       MPTester.Print "wait for MP Ready"
       Call LoadMP_Click_AU6988_K9F1G
 
        OldTimer = Timer
        AlcorMPMessage = 0
        Debug.Print "begin"
        Do
           ' DoEvents
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
            PassTime = Timer - OldTimer
            'Debug.Print AlcorMPMessage
        Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 30 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3
        '====================================================
        If PassTime > 30 Then
          '(1)
           MPTester.Print "MP Ready Fail"
            winHwnd = FindWindow(vbNullString, AU6988MPCaption1)
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, AU6988MPCaption1)
              Loop While winHwnd <> 0
            End If
           '(2)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
            If winHwnd <> 0 Then
              Do
              rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
              Call MsecDelay(0.5)
              winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
              Loop While winHwnd <> 0
            End If
            
        
             MPTester.TestResultLab = "Bin3:MP Ready Fail"
             TestResult = "Bin3"
             MPTester.Print "MP Ready Fail"
     
              
            Exit Sub
        End If
        
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_MP_START Then
         
             
             cardresult = DO_WritePort(card, Channel_P1A, &HFD)  ' sel chip
              Call PowerSet(500)   ' close power to disable chip
             
             
            Dim TimerCounter As Integer
            Dim TmpString As String
            
             
            Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName("vid")
            Loop While TmpString = "" And TimerCounter < 150
             
            Call MsecDelay(0.3)
             
             If TmpString = "" Then   ' can not find device after 15 s
             
               TestResult = "Bin2"
               MPTester.TestResultLab = "Bin2:MP UNKNOW Fail when enter MP"
               Exit Sub
             End If
             
             Call MsecDelay(2.5)
               
             MPTester.Print " MP Begin....."
             
             Call StartMP_Click_AU6988_K9F1G
   
              ReMP_Flag = 0
              OldTimer = Timer
              AlcorMPMessage = 0
                
                Do
                   'DoEvents
                    If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                        AlcorMPMessage = mMsg.message
                        TranslateMessage mMsg
                        DispatchMessage mMsg
                            
                        If (AlcorMPMessage = WM_FT_MP_FAIL) And (ReMP_Flag = 0) Then
                            AlcorMPMessage = 1
                            cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'close power
                            Call MsecDelay(0.3)
                            cardresult = DO_WritePort(card, Channel_P1A, &HFD)  'open power
                            Call MsecDelay(2.2)
                            Call RefreshMP_Click_AU6988
                            Call MsecDelay(0.5)
                            Call StartMP_Click_AU6988
                            
                            ReMP_Counter = ReMP_Counter + 1
                            If ReMP_Counter >= ReMP_Limit Then
                                ReMP_Flag = 1
                                ReMP_Counter = 0
                            End If
                        End If
                    End If
                    
                    PassTime = Timer - OldTimer
                
                Loop Until AlcorMPMessage = WM_FT_MP_PASS _
                Or AlcorMPMessage = WM_FT_MP_FAIL _
                Or AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL _
                Or PassTime > 65 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                
                If AlcorMPMessage = WM_FT_MP_PASS Then
                    If ReMP_Flag = 0 Then
                        MsecDelay (MPIdleTime * (ReMP_Limit - ReMP_Counter))
                    End If
                    ReMP_Counter = 0
                End If
                
                MPTester.Print "MP work time="; PassTime
                 MPTester.MPText.Text = Hex(AlcorMPMessage)
                '================================================
                '  Handle MP work time out error
                '===============================================
                
               ' time out fail
                If PassTime > 65 Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    '(1)
                    winHwnd = FindWindow(vbNullString, AU6988MPCaption)
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, AU6988MPCaption)
                      Loop While winHwnd <> 0
                    End If
                    
                    '(2)
                      winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    Exit Sub
                End If
                
                ' MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    
                    winHwnd = FindWindow(vbNullString, AU6988MPCaption)
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, AU6988MPCaption)
                      Loop While winHwnd <> 0
                    End If
                    
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    
                    Exit Sub
                End If
                
                
                 'unknow fail
                 If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                     MPContFail = MPContFail + 1
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                    MPTester.Print "MP UNKNOW Fail"
                    
                   winHwnd = FindWindow(vbNullString, AU6988MPCaption)
                    If winHwnd <> 0 Then
                     Do
                     rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                     Call MsecDelay(0.5)
                     winHwnd = FindWindow(vbNullString, AU6988MPCaption)
                     Loop While winHwnd <> 0
                     End If
                     
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                     If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                     End If
                     
                     
                     
                    Exit Sub
                End If
                 
                
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_PASS Then
                     MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPTester.Print "MP PASS"
                End If
        End If
   
End If
'=========================================
 '    Close MP program
 '=========================================
winHwnd = FindWindow(vbNullString, AU6988MPCaption)
If winHwnd <> 0 Then
  Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, AU6988MPCaption)
  Loop While winHwnd <> 0
    
    Call MsecDelay(0.2)
    KillProcess ("AlcorMP.exe")

End If
    
 winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
    Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
    Loop While winHwnd <> 0
End If

  Dim pid As Long          ' unload driver
  Dim hProcess As Long
  Dim ExitEvent As Long
 
  pid = Shell(App.Path & "\AlcorMP_6988\loaddrv.exe uninstall_058F6387")
  hProcess = OpenProcess(SYNCHRONIZE + PROCESS_QUERY_INFORMATION + PROCESS_TERMINATE, 0, pid)
  ExitEvent = WaitForSingleObject(hProcess, INFINITE)
  Call CloseHandle(hProcess)
  KillProcess ("LoadDrv.exe")
 
 
                        
 '=========================================
 '    POWER on
 '=========================================
'T1:
RW_Test_Label:
 
Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'Power OFF UNLoad Device
WaitDevOFF ("vid_058f")
Call MsecDelay(0.3)
 
 If MPFlag = 1 Then
         Call PowerSet(3)
          cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        
         Call MsecDelay(0.5)  ' power of to unload MPDriver

           cardresult = DO_WritePort(card, Channel_P1A, &HFA)  'sel socket
           'Call PowerSet(1500)
        Call PowerSet2(1, "5.0.", "0.18", 1, "5.0", "0.18", 1)
        
         Call MsecDelay(1.2)
        MPFlag = 0
 Else
          cardresult = DO_WritePort(card, Channel_P1A, &HFA)
         'Call PowerSet(1500)
         Call PowerSet2(1, "5.0.", "0.18", 1, "5.0", "0.18", 1)
         
         Call MsecDelay(1.2)
End If
         Call LoadRWTest_Click_AU6988_K9F1G

        
        
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
     '   GoTo T2
       If PassTime > 5 Then
           TestResult = "Bin3"
           MPTester.TestResultLab = "Bin3:RW Ready Fail"
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
       
            Exit Sub
       End If
         
T2:
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        Call StartRWTest_Click_AU6988
        
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
            Or AlcorMPMessage = WM_FT_CHECK_CERBGPO_FAIL _
            Or AlcorMPMessage = WM_FT_CHECK_HW_CODE_FAIL _
            Or AlcorMPMessage = WM_FT_PHYREAD_FAIL _
            Or AlcorMPMessage = WM_FT_ECC_FAIL _
            Or AlcorMPMessage = WM_FT_NOFREEBLOCK_FAIL _
            Or AlcorMPMessage = WM_FT_LODECODE_FAIL _
            Or AlcorMPMessage = WM_FT_RELOADCODE_FAIL _
            Or AlcorMPMessage = WM_FT_TESTUNITREADY_FAIL _
            Or PassTime > 10 _
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
                Exit Sub
            End If
        
        End If
               
               
        Select Case AlcorMPMessage
        
        Case WM_FT_RW_UNKNOW_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:UnKnow Fail"
             ContFail = ContFail + 1
        
        Case WM_FT_CHECK_HW_CODE_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:HW-ID Fail"
             ContFail = ContFail + 1
        
        Case WM_FT_TESTUNITREADY_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:TestUnitReady Fail"
             ContFail = ContFail + 1

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
        
        Case WM_FT_RELOADCODE_FAIL
              TestResult = "Bin4"
              MPTester.TestResultLab = "Bin4:ReLoadCode FAIL "
              ContFail = ContFail + 1
        
        Case WM_FT_ECC_FAIL
              TestResult = "Bin5"
              MPTester.TestResultLab = "Bin5:ECC FAIL "
              ContFail = ContFail + 1
                    
        Case WM_FT_RW_RW_PASS
        
               
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
              
               End If
               
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:Undefine Fail"
          
             ContFail = ContFail + 1
        
               
        End Select
                               
       cardresult = DO_WritePort(card, Channel_P1A, &HFA)
         'Call PowerSet(1500)
        Call PowerSet2(1, "5.0.", "0.18", 1, "5.0", "0.18", 1)
         
                            
                            
End Sub

Public Sub AU6988D53HLF20TestSub()

'Support K9F1G
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
 'Dim ReMP_Flag As Byte
 
 
 
   MPTester.TestResultLab = ""
'===============================================================
' Fail loacatio initial
'===============================================================
 
If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
    FileCopy App.Path & "\AlcorMP_6988\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
    Call MsecDelay(5)
End If


NewChipFlag = 0
If OldChipName <> ChipName Then
    FileCopy App.Path & "\AlcorMP_6988\INI\" & ChipName & "\ROM.Hex", App.Path & "\AlcorMP_6988\ROM.Hex"
    FileCopy App.Path & "\AlcorMP_6988\INI\" & ChipName & "\RAM.Bin", App.Path & "\AlcorMP_6988\RAM.Bin"
    FileCopy App.Path & "\AlcorMP_6988\INI\" & ChipName & "\AlcorMP.ini", App.Path & "\AlcorMP_6988\AlcorMP.ini"
    FileCopy App.Path & "\AlcorMP_6988\INI\" & ChipName & "\PE.bin", App.Path & "\AlcorMP_6988\PE.bin"
    FileCopy App.Path & "\AlcorMP_6988\INI\" & ChipName & "\FT.ini", App.Path & "\FT.ini"
    FileCopy App.Path & "\AlcorMP_6988\INI\" & ChipName & "\FT.ini", App.Path & "\AlcorMP_6988\FT.ini"
    NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName
 

'==============================================================
' when begin RW Test, must clear MP porgram
'===============================================================


'(1)  /// for Auto mode
winHwnd = FindWindow(vbNullString, AU6988MPCaption1)
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, AU6988MPCaption1)
  Loop While winHwnd <> 0
End If

'(1)
winHwnd = FindWindow(vbNullString, AU6988MPCaption)
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, AU6988MPCaption)
  Loop While winHwnd <> 0
End If
'(2)
winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
  Loop While winHwnd <> 0
End If
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail
 
 
 '====================================
 '  Fix Card
 '====================================
' GoTo T1
 If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag = True) Then
 
 
   If MPTester.NoMP.Value = 1 Then
        
        If (NewChipFlag = 0) And (MPTester.Check1.Value = 0) Then  ' force condition
           GoTo RW_Test_Label
        End If
    End If
       
    
    If MPTester.ResetMPFailCounter.Value = 1 Then
       ContFail = 0
    End If
    
 '==============================================================
' when begin MP, must close RW porgram
'===============================================================
   MPFlag = 1
 
    winHwnd = FindWindow(vbNullString, "UFD Test")
    If winHwnd <> 0 Then
      Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "UFD Test")
      Loop While winHwnd <> 0
    End If
 
       '  power on
       cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        Call PowerSet(3)   ' close power to disable chip
       Call MsecDelay(2.5)  ' power for load MPDriver
       MPTester.Print "wait for MP Ready"
       Call LoadMP_Click_AU6988_K9F1G
 
        OldTimer = Timer
        AlcorMPMessage = 0
        Debug.Print "begin"
        Do
           ' DoEvents
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
            PassTime = Timer - OldTimer
            'Debug.Print AlcorMPMessage
        Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 30 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3
        '====================================================
        If PassTime > 30 Then
          '(1)
           MPTester.Print "MP Ready Fail"
            winHwnd = FindWindow(vbNullString, AU6988MPCaption1)
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, AU6988MPCaption1)
              Loop While winHwnd <> 0
            End If
           '(2)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
            If winHwnd <> 0 Then
              Do
              rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
              Call MsecDelay(0.5)
              winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
              Loop While winHwnd <> 0
            End If
            
        
             MPTester.TestResultLab = "Bin3:MP Ready Fail"
             TestResult = "Bin3"
             MPTester.Print "MP Ready Fail"
     
              
            Exit Sub
        End If
        
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_MP_START Then
         
             
             cardresult = DO_WritePort(card, Channel_P1A, &HFB)  ' sel chip
              Call PowerSet(500)   ' close power to disable chip
             
             
            Dim TimerCounter As Integer
            Dim TmpString As String
            
             
            Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName("vid")
            Loop While TmpString = "" And TimerCounter < 150
             
            Call MsecDelay(0.3)
             
             If TmpString = "" Then   ' can not find device after 15 s
             
               TestResult = "Bin2"
               MPTester.TestResultLab = "Bin2:MP UNKNOW Fail when enter MP"
               Exit Sub
             End If
             
             Call MsecDelay(2.5)
               
             MPTester.Print " MP Begin....."
             
             Call StartMP_Click_AU6988_K9F1G
   
              ReMP_Flag = 0
              OldTimer = Timer
              AlcorMPMessage = 0
                
                Do
                   'DoEvents
                    If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                        AlcorMPMessage = mMsg.message
                        TranslateMessage mMsg
                        DispatchMessage mMsg
                            
                        If (AlcorMPMessage = WM_FT_MP_FAIL) And (ReMP_Flag = 0) Then
                            AlcorMPMessage = 1
                            cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'close power
                            Call MsecDelay(0.3)
                            cardresult = DO_WritePort(card, Channel_P1A, &HFB)  'open power
                            Call MsecDelay(2.2)
                            Call RefreshMP_Click_AU6988
                            Call MsecDelay(0.5)
                            Call StartMP_Click_AU6988
                            
                            ReMP_Counter = ReMP_Counter + 1
                            If ReMP_Counter >= ReMP_Limit Then
                                ReMP_Flag = 1
                                ReMP_Counter = 0
                            End If
                        End If
                    End If
                    
                    PassTime = Timer - OldTimer
                
                Loop Until AlcorMPMessage = WM_FT_MP_PASS _
                Or AlcorMPMessage = WM_FT_MP_FAIL _
                Or AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL _
                Or PassTime > 65 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                
                If AlcorMPMessage = WM_FT_MP_PASS Then
                    If ReMP_Flag = 0 Then
                        MsecDelay (MPIdleTime * (ReMP_Limit - ReMP_Counter))
                    End If
                    ReMP_Counter = 0
                End If
                
                MPTester.Print "MP work time="; PassTime
                 MPTester.MPText.Text = Hex(AlcorMPMessage)
                '================================================
                '  Handle MP work time out error
                '===============================================
                
               ' time out fail
                If PassTime > 65 Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    '(1)
                    winHwnd = FindWindow(vbNullString, AU6988MPCaption)
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, AU6988MPCaption)
                      Loop While winHwnd <> 0
                    End If
                    
                    '(2)
                      winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    Exit Sub
                End If
                
                ' MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    
                    winHwnd = FindWindow(vbNullString, AU6988MPCaption)
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, AU6988MPCaption)
                      Loop While winHwnd <> 0
                    End If
                    
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    
                    Exit Sub
                End If
                
                
                 'unknow fail
                 If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                     MPContFail = MPContFail + 1
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                    MPTester.Print "MP UNKNOW Fail"
                    
                   winHwnd = FindWindow(vbNullString, AU6988MPCaption)
                    If winHwnd <> 0 Then
                     Do
                     rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                     Call MsecDelay(0.5)
                     winHwnd = FindWindow(vbNullString, AU6988MPCaption)
                     Loop While winHwnd <> 0
                     End If
                     
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                     If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                     End If
                     
                     
                     
                    Exit Sub
                End If
                 
                
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_PASS Then
                     MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPTester.Print "MP PASS"
                End If
        End If
   
End If
'=========================================
 '    Close MP program
 '=========================================
winHwnd = FindWindow(vbNullString, AU6988MPCaption)
If winHwnd <> 0 Then
  Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, AU6988MPCaption)
  Loop While winHwnd <> 0
    
    Call MsecDelay(0.2)
    KillProcess ("AlcorMP.exe")

End If
    
 winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
    Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
    Loop While winHwnd <> 0
End If

  Dim pid As Long          ' unload driver
  Dim hProcess As Long
  Dim ExitEvent As Long
 
  pid = Shell(App.Path & "\AlcorMP_6988\loaddrv.exe uninstall_058F6387")
  hProcess = OpenProcess(SYNCHRONIZE + PROCESS_QUERY_INFORMATION + PROCESS_TERMINATE, 0, pid)
  ExitEvent = WaitForSingleObject(hProcess, INFINITE)
  Call CloseHandle(hProcess)
  KillProcess ("LoadDrv.exe")
 
 
                        
 '=========================================
 '    POWER on
 '=========================================
'T1:
RW_Test_Label:
 
Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'Power OFF UNLoad Device
WaitDevOFF ("vid_058f")
Call MsecDelay(0.3)
 
 If MPFlag = 1 Then
         Call PowerSet(3)
          cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        
         Call MsecDelay(0.5)  ' power of to unload MPDriver

           cardresult = DO_WritePort(card, Channel_P1A, &HFB)  'sel socket
           'Call PowerSet(1500)
        Call PowerSet2(1, "5.0.", "0.18", 1, "5.0", "0.18", 1)
        
         'Call MsecDelay(1.2)
         WaitDevOn ("vid_058f")
        MPFlag = 0
 Else
          cardresult = DO_WritePort(card, Channel_P1A, &HFA)
         'Call PowerSet(1500)
         Call PowerSet2(1, "5.0.", "0.18", 1, "5.0", "0.18", 1)
         
         'Call MsecDelay(1.2)
         WaitDevOn ("vid_058f")
End If
         Call LoadRWTest_Click_AU6988_K9F1G

        
        
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
     '   GoTo T2
       If PassTime > 5 Then
           TestResult = "Bin3"
           MPTester.TestResultLab = "Bin3:RW Ready Fail"
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
       
            Exit Sub
       End If
         
T2:
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        Call StartRWTest_Click_AU6988
        
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
            Or AlcorMPMessage = WM_FT_CHECK_CERBGPO_FAIL _
            Or AlcorMPMessage = WM_FT_CHECK_HW_CODE_FAIL _
            Or AlcorMPMessage = WM_FT_PHYREAD_FAIL _
            Or AlcorMPMessage = WM_FT_ECC_FAIL _
            Or AlcorMPMessage = WM_FT_NOFREEBLOCK_FAIL _
            Or AlcorMPMessage = WM_FT_LODECODE_FAIL _
            Or AlcorMPMessage = WM_FT_RELOADCODE_FAIL _
            Or AlcorMPMessage = WM_FT_TESTUNITREADY_FAIL _
            Or PassTime > 10 _
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
                Exit Sub
            End If
        
        End If
               
               
        Select Case AlcorMPMessage
        
        Case WM_FT_RW_UNKNOW_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:UnKnow Fail"
             ContFail = ContFail + 1
        
        Case WM_FT_CHECK_HW_CODE_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:HW-ID Fail"
             ContFail = ContFail + 1
        
        Case WM_FT_TESTUNITREADY_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:TestUnitReady Fail"
             ContFail = ContFail + 1

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
        
        Case WM_FT_RELOADCODE_FAIL
              TestResult = "Bin4"
              MPTester.TestResultLab = "Bin4:ReLoadCode FAIL "
              ContFail = ContFail + 1
        
        Case WM_FT_ECC_FAIL
              TestResult = "Bin5"
              MPTester.TestResultLab = "Bin5:ECC FAIL "
              ContFail = ContFail + 1
                    
        Case WM_FT_RW_RW_PASS
        
               
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
              
               End If
               
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:Undefine Fail"
          
             ContFail = ContFail + 1
        
               
        End Select
                               
       cardresult = DO_WritePort(card, Channel_P1A, &HFF)
         'Call PowerSet(1500)
        Call PowerSet2(1, "0.0.", "0.18", 1, "0.0", "0.18", 1)
         
                            
                            
End Sub

Public Sub LoadMP_Click_AU6988_K9F1G()

Dim TimePass
Dim rt2
' find window
 
    winHwnd = FindWindow(vbNullString, "Module Update")
 
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\AlcorMP_6988\AU699x_MP_Update_Patch_v11.09.26.01-FT.exe", "", "", SW_SHOW)
    End If



    Do
        
        winHwnd = FindWindow(vbNullString, "Module Update")
    
    Loop While winHwnd = 0
 
    winHwnd = FindWindow(vbNullString, "Module Update")
    
    Do
        
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        winHwnd = FindWindow(vbNullString, "Module Update")
    
    Loop While winHwnd <> 0
 
 
 
    winHwnd = FindWindow(vbNullString, AU6988MPCaption1)
 
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\AlcorMP_6988\AlcorMP.exe", "", "", SW_SHOW)
    End If

    SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags

End Sub

Public Sub StartMP_Click_AU6988_K9F1G()
Dim rt2
    winHwnd = FindWindow(vbNullString, AU6988MPCaption)
    Debug.Print "WindHandle="; winHwnd
    rt2 = PostMessage(winHwnd, WM_FT_MP_START, 0&, 0&)
End Sub

 Public Sub RefreshMP_Click_AU6988()
 Dim rt2
    winHwnd = FindWindow(vbNullString, AU6988MPCaption)
    Debug.Print "WindHandle="; winHwnd
    rt2 = PostMessage(winHwnd, WM_FT_MP_REFRESH, 0&, 0&)
 End Sub

Public Sub AU6988D52HLF29TestSub()
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
 
If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
    FileCopy App.Path & "\AU6986AU6988_20090904\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
    Call MsecDelay(5)
End If


NewChipFlag = 0
If OldChipName <> ChipName Then
            FileCopy App.Path & "\AU6986AU6988_20090904\ROM\" & ChipName & "\ROM.Hex", App.Path & "\AU6986AU6988_20090904\ROM.Hex"
            FileCopy App.Path & "\AU6986AU6988_20090904\RAM\" & ChipName & "\RAM.Bin", App.Path & "\AU6986AU6988_20090904\RAM.Bin"
            FileCopy App.Path & "\AU6986AU6988_20090904\INI\" & ChipName & "\AlcorMP.ini", App.Path & "\AU6986AU6988_20090904\AlcorMP.ini"
            NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName
 

'==============================================================
' when begin RW Test, must clear MP program
'===============================================================
'(1)
winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
  Loop While winHwnd <> 0
End If
'(2)
winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
  Loop While winHwnd <> 0
End If
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail
 
 
 '====================================
 '  Fix Card
 '====================================
 
 If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag = True) Then
    
    If MPTester.NoMP.Value = 1 Then
        
        If (NewChipFlag = 0) And (MPTester.Check1.Value = 0) Then  ' force condition
           GoTo RW_Test_Label
        End If
    End If
       
    
    If MPTester.ResetMPFailCounter.Value = 1 Then
       ContFail = 0
    End If
    

    
 '==============================================================
' when begin MP, must close RW program
'===============================================================
   MPFlag = 1
 
    winHwnd = FindWindow(vbNullString, "UFD Test")
    If winHwnd <> 0 Then
      Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "UFD Test")
      Loop While winHwnd <> 0
    End If
 
       '  power on
       cardresult = DO_WritePort(card, Channel_P1A, &HFF)
       Call PowerSet(3)   ' close power to disable chip
       Call MsecDelay(0.5)  ' power for load MPDriver
       MPTester.Print "wait for MP Ready"
       Call LoadMP_Click_AU6988_20090904
 
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
        Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 30 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3
        '====================================================
        If PassTime > 30 Then
          '(1)
           MPTester.Print "MP Ready Fail"
            winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
              Loop While winHwnd <> 0
            End If
           '(2)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
            If winHwnd <> 0 Then
              Do
              rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
              Call MsecDelay(0.5)
              winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
              Loop While winHwnd <> 0
            End If
            
        
             MPTester.TestResultLab = "Bin3:MP Ready Fail"
             TestResult = "Bin3"
             MPTester.Print "MP Ready Fail"
     
              
            Exit Sub
        End If
        
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_MP_START Then
         
             
             cardresult = DO_WritePort(card, Channel_P1A, &HFD)  ' sel chip
             Call PowerSet(500)   ' close power to disable chip
             
 
           '  Call MsecDelay(6.5)
             
            Dim TimerCounter As Integer
            Dim TmpString As String
            
             
            Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName("vid")
            Loop While TmpString = "" And TimerCounter < 150
             
             
             If TmpString = "" Then   ' can not find device after 15 s
             
               TestResult = "Bin2"
               MPTester.TestResultLab = "Bin2:MP UNKNOW Fail when enter MP"
               Exit Sub
             End If
             
             Call MsecDelay(2.5)
               
             MPTester.Print " MP Begin....."
             
             Call StartMP_Click_AU6988
   
             
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
                
               ' time out fail
                If PassTime > 65 Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    '(1)
                    winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                      Loop While winHwnd <> 0
                    End If
                    
                    '(2)
                      winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    Exit Sub
                End If
                
                ' MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    
                    winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                      Loop While winHwnd <> 0
                    End If
                    
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    
                    Exit Sub
                End If
                
                
                 'unknow fail
                 If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                     MPContFail = MPContFail + 1
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                    MPTester.Print "MP UNKNOW Fail"
                    
                   winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                     Do
                     rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                     Call MsecDelay(0.5)
                     winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                     Loop While winHwnd <> 0
                     End If
                     
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                     If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                     End If
                     
                     
                     
                    Exit Sub
                End If
                 
                
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_PASS Then
                     MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPTester.Print "MP PASS"
                End If
        End If
   
End If
'=========================================
 '    Close MP program
 '=========================================
winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
If winHwnd <> 0 Then
  Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
  Loop While winHwnd <> 0
End If
    
 winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
    Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
    Loop While winHwnd <> 0
End If

  Dim pid As Long          ' unload driver
  Dim hProcess As Long
  Dim ExitEvent As Long
 
  pid = Shell(App.Path & "\AU6986AU6988_20090904\loaddrv.exe uninstall_058F6387")
  hProcess = OpenProcess(SYNCHRONIZE + PROCESS_QUERY_INFORMATION + PROCESS_TERMINATE, 0, pid)
  ExitEvent = WaitForSingleObject(hProcess, INFINITE)
  Call CloseHandle(hProcess)
 
 
 
                        
 '=========================================
 '    POWER on
 '=========================================
RW_Test_Label:
 If MPFlag = 1 Then
        Call PowerSet(3)
          cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        
         Call MsecDelay(0.5)  ' power of to unload MPDriver

           cardresult = DO_WritePort(card, Channel_P1A, &HF2)  'sel socket
          Call PowerSet(1900)
     
        
         Call MsecDelay(1.2)
        MPFlag = 0
 Else
          cardresult = DO_WritePort(card, Channel_P1A, &HF2)
         Call PowerSet(1900)
         
         Call MsecDelay(1.2)
End If
         Call LoadRWTest_Click_AU6988_20090904

        
        
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
        
       If PassTime > 5 Then
           TestResult = "Bin3"
           MPTester.TestResultLab = "Bin3:RW Ready Fail"
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
       
            Exit Sub
       End If
         
         
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        Call StartRWTest_Click_AU6988
        
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
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
            
              cardresult = DO_WritePort(card, Channel_P1A, &HF2)  ' power off
            
            Call PowerSet(1900)
        
            
       
            Exit Sub
        End If
        
        
     
               
        Select Case AlcorMPMessage
        
        Case WM_FT_RW_UNKNOW_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:UnKnow Fail"
          
             ContFail = ContFail + 1
        
        Case WM_FT_RW_SPEED_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:SPEED Error "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_RW_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:RW FAIL "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_ROM_FAIL
              TestResult = "Bin4"
              MPTester.TestResultLab = "Bin4:ROM FAIL "
              ContFail = ContFail + 1
              
        Case WM_FT_RW_RAM_FAIL
              TestResult = "Bin5"
              MPTester.TestResultLab = "Bin5:RAM FAIL "
               ContFail = ContFail + 1
        Case WM_FT_RW_RW_PASS
        
               
               For LedCount = 1 To 20
               Call MsecDelay(0.1)
               cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
                If LightOn = &HEF Or LightOn = &HCF Then
               
                 Exit For
               End If
               Next LedCount
                 
                  MPTester.Print "light="; LightOn
                 If LightOn = &HEF Or LightOn = &HCF Then
                    MPTester.TestResultLab = "PASS "
                    TestResult = "PASS"
                    ContFail = 0
                Else
                 
                  TestResult = "Bin3"
                  MPTester.TestResultLab = "Bin3:LED FAIL "
              
               End If
               
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:Undefine Fail"
          
             ContFail = ContFail + 1
        
               
        End Select
                               
       cardresult = DO_WritePort(card, Channel_P1A, &HF2)
        Call PowerSet(1900)
        
         
                            
End Sub
Public Sub AU6988D52HLF27TestSub()
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
 
If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
    FileCopy App.Path & "\AlcorMP_698x_PD\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
    Call MsecDelay(5)
End If


NewChipFlag = 0
If OldChipName <> ChipName Then
            FileCopy App.Path & "\AlcorMP_698x_PD\ROM\" & ChipName & "\ROM.Hex", App.Path & "\AlcorMP_698x_PD\ROM.Hex"
            FileCopy App.Path & "\AlcorMP_698x_PD\RAM\" & ChipName & "\RAM.Bin", App.Path & "\AlcorMP_698x_PD\RAM.Bin"
            FileCopy App.Path & "\AlcorMP_698x_PD\INI\" & ChipName & "\AlcorMP.ini", App.Path & "\AlcorMP_698x_PD\AlcorMP.ini"
            NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName
 

'==============================================================
' when begin RW Test, must clear MP program
'===============================================================
'(1)
winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
  Loop While winHwnd <> 0
End If
'(2)
winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
  Loop While winHwnd <> 0
End If
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail
 
 
 '====================================
 '  Fix Card
 '====================================
 
 If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag = True) Then
    
    If MPTester.NoMP.Value = 1 Then
        
        If (NewChipFlag = 0) And (MPTester.Check1.Value = 0) Then  ' force condition
           GoTo RW_Test_Label
        End If
    End If
       
    
    If MPTester.ResetMPFailCounter.Value = 1 Then
       ContFail = 0
    End If
    

    
 '==============================================================
' when begin MP, must close RW program
'===============================================================
   MPFlag = 1
 
    winHwnd = FindWindow(vbNullString, "UFD Test")
    If winHwnd <> 0 Then
      Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "UFD Test")
      Loop While winHwnd <> 0
    End If
 
       '  power on
       cardresult = DO_WritePort(card, Channel_P1A, &HFF)
       Call PowerSet(3)   ' close power to disable chip
       Call MsecDelay(0.5)  ' power for load MPDriver
       MPTester.Print "wait for MP Ready"
       Call LoadMP_Click_AU6988
 
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
        Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 30 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3
        '====================================================
        If PassTime > 30 Then
          '(1)
           MPTester.Print "MP Ready Fail"
            winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
              Loop While winHwnd <> 0
            End If
           '(2)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
            If winHwnd <> 0 Then
              Do
              rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
              Call MsecDelay(0.5)
              winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
              Loop While winHwnd <> 0
            End If
            
        
             MPTester.TestResultLab = "Bin3:MP Ready Fail"
             TestResult = "Bin3"
             MPTester.Print "MP Ready Fail"
     
              
            Exit Sub
        End If
        
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_MP_START Then
         
             
             cardresult = DO_WritePort(card, Channel_P1A, &HFD)  ' sel chip
             Call PowerSet(500)   ' close power to disable chip
             
 
           '  Call MsecDelay(6.5)
             
            Dim TimerCounter As Integer
            Dim TmpString As String
            
             
            Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName("vid")
            Loop While TmpString = "" And TimerCounter < 150
             
             
             If TmpString = "" Then   ' can not find device after 15 s
             
               TestResult = "Bin2"
               MPTester.TestResultLab = "Bin2:MP UNKNOW Fail when enter MP"
               Exit Sub
             End If
             
             Call MsecDelay(2.5)
               
             MPTester.Print " MP Begin....."
             
             Call StartMP_Click_AU6988
   
             
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
                
               ' time out fail
                If PassTime > 65 Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    '(1)
                    winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                      Loop While winHwnd <> 0
                    End If
                    
                    '(2)
                      winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    Exit Sub
                End If
                
                ' MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    
                    winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                      Loop While winHwnd <> 0
                    End If
                    
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    
                    Exit Sub
                End If
                
                
                 'unknow fail
                 If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                     MPContFail = MPContFail + 1
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                    MPTester.Print "MP UNKNOW Fail"
                    
                   winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                     Do
                     rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                     Call MsecDelay(0.5)
                     winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                     Loop While winHwnd <> 0
                     End If
                     
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                     If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                     End If
                     
                     
                     
                    Exit Sub
                End If
                 
                
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_PASS Then
                     MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPTester.Print "MP PASS"
                End If
        End If
   
End If
'=========================================
 '    Close MP program
 '=========================================
winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
If winHwnd <> 0 Then
  Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
  Loop While winHwnd <> 0
End If
    
 winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
    Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
    Loop While winHwnd <> 0
End If

  Dim pid As Long          ' unload driver
  Dim hProcess As Long
  Dim ExitEvent As Long
 
  pid = Shell(App.Path & "\AlcorMP_698x_PD\loaddrv.exe uninstall_058F6387")
  hProcess = OpenProcess(SYNCHRONIZE + PROCESS_QUERY_INFORMATION + PROCESS_TERMINATE, 0, pid)
  ExitEvent = WaitForSingleObject(hProcess, INFINITE)
  Call CloseHandle(hProcess)
 
 
 
                        
 '=========================================
 '    POWER on
 '=========================================
RW_Test_Label:
 If MPFlag = 1 Then
        Call PowerSet(3)
          cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        
         Call MsecDelay(0.5)  ' power of to unload MPDriver

           cardresult = DO_WritePort(card, Channel_P1A, &HF2)  'sel socket
          Call PowerSet(1500)
     
        
         Call MsecDelay(1.2)
        MPFlag = 0
 Else
          cardresult = DO_WritePort(card, Channel_P1A, &HF2)
         Call PowerSet(1500)
         
         Call MsecDelay(1.2)
End If
         Call LoadRWTest_Click_AU6988

        
        
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
        
       If PassTime > 5 Then
           TestResult = "Bin3"
           MPTester.TestResultLab = "Bin3:RW Ready Fail"
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
       
            Exit Sub
       End If
         
         
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        Call StartRWTest_Click_AU6988
        
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
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
            
              cardresult = DO_WritePort(card, Channel_P1A, &HF2)  ' power off
            
            Call PowerSet(1500)
        
            
       
            Exit Sub
        End If
        
        
     
               
        Select Case AlcorMPMessage
        
        Case WM_FT_RW_UNKNOW_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:UnKnow Fail"
          
             ContFail = ContFail + 1
        
        Case WM_FT_RW_SPEED_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:SPEED Error "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_RW_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:RW FAIL "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_ROM_FAIL
              TestResult = "Bin4"
              MPTester.TestResultLab = "Bin4:ROM FAIL "
              ContFail = ContFail + 1
              
        Case WM_FT_RW_RAM_FAIL
              TestResult = "Bin5"
              MPTester.TestResultLab = "Bin5:RAM FAIL "
               ContFail = ContFail + 1
        Case WM_FT_RW_RW_PASS
        
               
               For LedCount = 1 To 20
               Call MsecDelay(0.1)
               cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
                If LightOn = &HEF Or LightOn = &HCF Then
               
                 Exit For
               End If
               Next LedCount
                 
                  MPTester.Print "light="; LightOn
                 If LightOn = &HEF Or LightOn = &HCF Then
                    MPTester.TestResultLab = "PASS "
                    TestResult = "PASS"
                    ContFail = 0
                Else
                 
                  TestResult = "Bin3"
                  MPTester.TestResultLab = "Bin3:LED FAIL "
              
               End If
               
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:Undefine Fail"
          
             ContFail = ContFail + 1
        
               
        End Select
                               
       cardresult = DO_WritePort(card, Channel_P1A, &HF2)
        Call PowerSet(1500)
        
         
                            
End Sub
Public Sub AU6988H56ILF2ATestSub()
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
 
If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
    FileCopy App.Path & "\AU6986AU6988_20090904\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
    Call MsecDelay(5)
End If


NewChipFlag = 0
If OldChipName <> ChipName Then
            FileCopy App.Path & "\AU6986AU6988_20090904\ROM\" & ChipName & "\ROM.Hex", App.Path & "\AU6986AU6988_20090904\ROM.Hex"
            FileCopy App.Path & "\AU6986AU6988_20090904\RAM\" & ChipName & "\RAM.Bin", App.Path & "\AU6986AU6988_20090904\RAM.Bin"
            FileCopy App.Path & "\AU6986AU6988_20090904\INI\" & ChipName & "\AlcorMP.ini", App.Path & "\AU6986AU6988_20090904\AlcorMP.ini"
            NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName
 

'==============================================================
' when begin RW Test, must clear MP program
'===============================================================
'(1)
winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
  Loop While winHwnd <> 0
End If
'(2)
winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
  Loop While winHwnd <> 0
End If
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail
 
 
 '====================================
 '  Fix Card
 '====================================
 
 If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag = True) Then
 
  If MPTester.NoMP.Value = 1 Then
        
        If (NewChipFlag = 0) And (MPTester.Check1.Value = 0) Then  ' force condition
           GoTo RW_Test_Label
        End If
    End If
       
    
    If MPTester.ResetMPFailCounter.Value = 1 Then
       ContFail = 0
    End If
    
 '==============================================================
' when begin MP, must close RW program
'===============================================================
   MPFlag = 1
 
    winHwnd = FindWindow(vbNullString, "UFD Test")
    If winHwnd <> 0 Then
      Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "UFD Test")
      Loop While winHwnd <> 0
    End If
 
       '  power on
       cardresult = DO_WritePort(card, Channel_P1A, &HFF)
       Call PowerSet(3)   ' close power to disable chip
       
       Call MsecDelay(0.5)  ' power for load MPDriver
       MPTester.Print "wait for MP Ready"
       Call LoadMP_Click_AU6988_20090904
 
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
        Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 30 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3
        '====================================================
        If PassTime > 30 Then
          '(1)
           MPTester.Print "MP Ready Fail"
            winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
              Loop While winHwnd <> 0
            End If
           '(2)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
            If winHwnd <> 0 Then
              Do
              rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
              Call MsecDelay(0.5)
              winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
              Loop While winHwnd <> 0
            End If
            
        
             MPTester.TestResultLab = "Bin3:MP Ready Fail"
             TestResult = "Bin3"
             MPTester.Print "MP Ready Fail"
     
              
            Exit Sub
        End If
        
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_MP_START Then
         
             
             cardresult = DO_WritePort(card, Channel_P1A, &H0)  ' sel chip
             Call PowerSet(500)   ' close power to disable chip
             
 
           '  Call MsecDelay(6.5)
             
            Dim TimerCounter As Integer
            Dim TmpString As String
            
             
            Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName("vid")
            Loop While TmpString = "" And TimerCounter < 150
             
             
             If TmpString = "" Then   ' can not find device after 15 s
             
               TestResult = "Bin2"
               MPTester.TestResultLab = "Bin2:MP UNKNOW Fail when enter MP"
               Exit Sub
             End If
             
             Call MsecDelay(2.5)
               
             MPTester.Print " MP Begin....."
             
             Call StartMP_Click_AU6988
   
             
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
                
               ' time out fail
                If PassTime > 65 Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    '(1)
                    winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                      Loop While winHwnd <> 0
                    End If
                    
                    '(2)
                      winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    Exit Sub
                End If
                
                ' MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    
                    winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                      Loop While winHwnd <> 0
                    End If
                    
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    
                    Exit Sub
                End If
                
                
                 'unknow fail
                 If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                     MPContFail = MPContFail + 1
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                    MPTester.Print "MP UNKNOW Fail"
                    
                   winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                     Do
                     rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                     Call MsecDelay(0.5)
                     winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                     Loop While winHwnd <> 0
                     End If
                     
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                     If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                     End If
                     
                     
                     
                    Exit Sub
                End If
                 
                
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_PASS Then
                     MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPTester.Print "MP PASS"
                End If
        End If
   
End If
'=========================================
 '    Close MP program
 '=========================================
winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
If winHwnd <> 0 Then
  Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
  Loop While winHwnd <> 0
End If
    
 winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
    Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
    Loop While winHwnd <> 0
End If

  Dim pid As Long          ' unload driver
  Dim hProcess As Long
  Dim ExitEvent As Long
 
  pid = Shell(App.Path & "\AU6986AU6988_20090904\loaddrv.exe uninstall_058F6387")
  hProcess = OpenProcess(SYNCHRONIZE + PROCESS_QUERY_INFORMATION + PROCESS_TERMINATE, 0, pid)
  ExitEvent = WaitForSingleObject(hProcess, INFINITE)
  Call CloseHandle(hProcess)
 
 
 
                        
 '=========================================
 '    POWER on
 '=========================================
RW_Test_Label:
 If MPFlag = 1 Then
        Call PowerSet(3)
          cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        
         Call MsecDelay(0.5)  ' power of to unload MPDriver

           cardresult = DO_WritePort(card, Channel_P1A, &H0)  'sel socket
          Call PowerSet(1900)
     
        
         Call MsecDelay(1.2)
        MPFlag = 0
 Else
          cardresult = DO_WritePort(card, Channel_P1A, &H0)
         Call PowerSet(1900)
         
         Call MsecDelay(1.2)
End If
         Call LoadRWTest_Click_AU6988_20090904

        
        
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
        
       If PassTime > 5 Then
           TestResult = "Bin3"
           MPTester.TestResultLab = "Bin3:RW Ready Fail"
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
       
            Exit Sub
       End If
         
         
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        Call StartRWTest_Click_AU6988
        
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
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
            
              cardresult = DO_WritePort(card, Channel_P1A, &HFB)  ' power off
            
            Call PowerSet(1900)
        
            
       
            Exit Sub
        End If
        
        
     
               
        Select Case AlcorMPMessage
        
        Case WM_FT_RW_UNKNOW_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:UnKnow Fail"
          
             ContFail = ContFail + 1
        
        Case WM_FT_RW_SPEED_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:SPEED Error "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_RW_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:RW FAIL "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_ROM_FAIL
              TestResult = "Bin4"
              MPTester.TestResultLab = "Bin4:ROM FAIL "
              ContFail = ContFail + 1
              
        Case WM_FT_RW_RAM_FAIL
              TestResult = "Bin5"
              MPTester.TestResultLab = "Bin5:RAM FAIL "
               ContFail = ContFail + 1
        Case WM_FT_RW_RW_PASS
        
               
               For LedCount = 1 To 20
               Call MsecDelay(0.1)
               cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
                If LightOn = 254 Then
               
                 Exit For
               End If
               Next LedCount
                 
                  MPTester.Print "light="; LightOn
                 If LightOn = 254 Then
                    MPTester.TestResultLab = "PASS "
                    TestResult = "PASS"
                    ContFail = 0
                Else
                 
                  TestResult = "Bin3"
                  MPTester.TestResultLab = "Bin3:LED FAIL "
              
               End If
               
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:Undefine Fail"
          
             ContFail = ContFail + 1
        
               
        End Select
                               
       cardresult = DO_WritePort(card, Channel_P1A, &HFB)
        Call PowerSet(1900)
        
         
                            
End Sub

Public Sub AU6988H55ILF28TestSub()
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
 
If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
    FileCopy App.Path & "\AlcorMP_698x_PD\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
    Call MsecDelay(5)
End If


NewChipFlag = 0
If OldChipName <> ChipName Then
            FileCopy App.Path & "\AlcorMP_698x_PD\ROM\" & ChipName & "\ROM.Hex", App.Path & "\AlcorMP_698x_PD\ROM.Hex"
            FileCopy App.Path & "\AlcorMP_698x_PD\RAM\" & ChipName & "\RAM.Bin", App.Path & "\AlcorMP_698x_PD\RAM.Bin"
            FileCopy App.Path & "\AlcorMP_698x_PD\INI\" & ChipName & "\AlcorMP.ini", App.Path & "\AlcorMP_698x_PD\AlcorMP.ini"
            NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName
 

'==============================================================
' when begin RW Test, must clear MP program
'===============================================================
'(1)
winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
  Loop While winHwnd <> 0
End If
'(2)
winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
  Loop While winHwnd <> 0
End If
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail
 
 
 '====================================
 '  Fix Card
 '====================================
 
 If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag = True) Then
    
 '==============================================================
' when begin MP, must close RW program
'===============================================================
   MPFlag = 1
 
    winHwnd = FindWindow(vbNullString, "UFD Test")
    If winHwnd <> 0 Then
      Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "UFD Test")
      Loop While winHwnd <> 0
    End If
 
       '  power on
       cardresult = DO_WritePort(card, Channel_P1A, &HFF)
       Call PowerSet(3)   ' close power to disable chip
       
       Call MsecDelay(0.5)  ' power for load MPDriver
       MPTester.Print "wait for MP Ready"
       Call LoadMP_Click_AU6988
 
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
        Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 30 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3
        '====================================================
        If PassTime > 30 Then
          '(1)
           MPTester.Print "MP Ready Fail"
            winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
              Loop While winHwnd <> 0
            End If
           '(2)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
            If winHwnd <> 0 Then
              Do
              rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
              Call MsecDelay(0.5)
              winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
              Loop While winHwnd <> 0
            End If
            
        
             MPTester.TestResultLab = "Bin3:MP Ready Fail"
             TestResult = "Bin3"
             MPTester.Print "MP Ready Fail"
     
              
            Exit Sub
        End If
        
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_MP_START Then
         
             
             cardresult = DO_WritePort(card, Channel_P1A, &H0)  ' sel chip
             Call PowerSet(500)   ' close power to disable chip
             
 
           '  Call MsecDelay(6.5)
             
            Dim TimerCounter As Integer
            Dim TmpString As String
            
             
            Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName("vid")
            Loop While TmpString = "" And TimerCounter < 150
             
             
             If TmpString = "" Then   ' can not find device after 15 s
             
               TestResult = "Bin2"
               MPTester.TestResultLab = "Bin2:MP UNKNOW Fail when enter MP"
               Exit Sub
             End If
             
             Call MsecDelay(2.5)
               
             MPTester.Print " MP Begin....."
             
             Call StartMP_Click_AU6988
   
             
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
                
               ' time out fail
                If PassTime > 65 Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    '(1)
                    winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                      Loop While winHwnd <> 0
                    End If
                    
                    '(2)
                      winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    Exit Sub
                End If
                
                ' MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    
                    winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                      Loop While winHwnd <> 0
                    End If
                    
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    
                    Exit Sub
                End If
                
                
                 'unknow fail
                 If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                     MPContFail = MPContFail + 1
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                    MPTester.Print "MP UNKNOW Fail"
                    
                   winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                     Do
                     rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                     Call MsecDelay(0.5)
                     winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                     Loop While winHwnd <> 0
                     End If
                     
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                     If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                     End If
                     
                     
                     
                    Exit Sub
                End If
                 
                
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_PASS Then
                     MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPTester.Print "MP PASS"
                End If
        End If
   
End If
'=========================================
 '    Close MP program
 '=========================================
winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
If winHwnd <> 0 Then
  Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
  Loop While winHwnd <> 0
End If
    
 winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
    Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
    Loop While winHwnd <> 0
End If

  Dim pid As Long          ' unload driver
  Dim hProcess As Long
  Dim ExitEvent As Long
 
  pid = Shell(App.Path & "\AlcorMP_698x_PD\loaddrv.exe uninstall_058F6387")
  hProcess = OpenProcess(SYNCHRONIZE + PROCESS_QUERY_INFORMATION + PROCESS_TERMINATE, 0, pid)
  ExitEvent = WaitForSingleObject(hProcess, INFINITE)
  Call CloseHandle(hProcess)
 
 
 
                        
 '=========================================
 '    POWER on
 '=========================================
 
 If MPFlag = 1 Then
        Call PowerSet(3)
          cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        
         Call MsecDelay(0.5)  ' power of to unload MPDriver

           cardresult = DO_WritePort(card, Channel_P1A, &H0)  'sel socket
          Call PowerSet(1900)
     
        
         Call MsecDelay(1.2)
        MPFlag = 0
 Else
          cardresult = DO_WritePort(card, Channel_P1A, &H0)
         Call PowerSet(1900)
         
         Call MsecDelay(1.2)
End If
         Call LoadRWTest_Click_AU6988

        
        
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
        
       If PassTime > 5 Then
           TestResult = "Bin3"
           MPTester.TestResultLab = "Bin3:RW Ready Fail"
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
       
            Exit Sub
       End If
         
         
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        Call StartRWTest_Click_AU6988
        
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
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
            
              cardresult = DO_WritePort(card, Channel_P1A, &HFB)  ' power off
            
            Call PowerSet(1900)
        
            
       
            Exit Sub
        End If
        
        
     
               
        Select Case AlcorMPMessage
        
        Case WM_FT_RW_UNKNOW_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:UnKnow Fail"
          
             ContFail = ContFail + 1
        
        Case WM_FT_RW_SPEED_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:SPEED Error "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_RW_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:RW FAIL "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_ROM_FAIL
              TestResult = "Bin4"
              MPTester.TestResultLab = "Bin4:ROM FAIL "
              ContFail = ContFail + 1
              
        Case WM_FT_RW_RAM_FAIL
              TestResult = "Bin5"
              MPTester.TestResultLab = "Bin5:RAM FAIL "
               ContFail = ContFail + 1
        Case WM_FT_RW_RW_PASS
        
               
               For LedCount = 1 To 20
               Call MsecDelay(0.1)
               cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
                If LightOn = 254 Then
               
                 Exit For
               End If
               Next LedCount
                 
                  MPTester.Print "light="; LightOn
                 If LightOn = 254 Then
                    MPTester.TestResultLab = "PASS "
                    TestResult = "PASS"
                    ContFail = 0
                Else
                 
                  TestResult = "Bin3"
                  MPTester.TestResultLab = "Bin3:LED FAIL "
              
               End If
               
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:Undefine Fail"
          
             ContFail = ContFail + 1
        
               
        End Select
                               
       cardresult = DO_WritePort(card, Channel_P1A, &HFB)
        Call PowerSet(1900)
        
         
                            
End Sub

Public Sub AU6988H55ILF26TestSub()
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
 
If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
    FileCopy App.Path & "\AlcorMP_698x_PD\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
    Call MsecDelay(5)
End If


NewChipFlag = 0
If OldChipName <> ChipName Then
            FileCopy App.Path & "\AlcorMP_698x_PD\ROM\" & ChipName & "\ROM.Hex", App.Path & "\AlcorMP_698x_PD\ROM.Hex"
            FileCopy App.Path & "\AlcorMP_698x_PD\RAM\" & ChipName & "\RAM.Bin", App.Path & "\AlcorMP_698x_PD\RAM.Bin"
            FileCopy App.Path & "\AlcorMP_698x_PD\INI\" & ChipName & "\AlcorMP.ini", App.Path & "\AlcorMP_698x_PD\AlcorMP.ini"
            NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName
 

'==============================================================
' when begin RW Test, must clear MP program
'===============================================================
'(1)
winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
  Loop While winHwnd <> 0
End If
'(2)
winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
  Loop While winHwnd <> 0
End If
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail
 
 
 '====================================
 '  Fix Card
 '====================================
 
 If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag = True) Then
    
 '==============================================================
' when begin MP, must close RW program
'===============================================================
   MPFlag = 1
 
    winHwnd = FindWindow(vbNullString, "UFD Test")
    If winHwnd <> 0 Then
      Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "UFD Test")
      Loop While winHwnd <> 0
    End If
 
       '  power on
       cardresult = DO_WritePort(card, Channel_P1A, &HFF)
       Call PowerSet(3)   ' close power to disable chip
       Call MsecDelay(0.5)  ' power for load MPDriver
       MPTester.Print "wait for MP Ready"
       Call LoadMP_Click_AU6988
 
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
        Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 30 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3
        '====================================================
        If PassTime > 30 Then
          '(1)
           MPTester.Print "MP Ready Fail"
            winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
              Loop While winHwnd <> 0
            End If
           '(2)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
            If winHwnd <> 0 Then
              Do
              rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
              Call MsecDelay(0.5)
              winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
              Loop While winHwnd <> 0
            End If
            
        
             MPTester.TestResultLab = "Bin3:MP Ready Fail"
             TestResult = "Bin3"
             MPTester.Print "MP Ready Fail"
     
              
            Exit Sub
        End If
        
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_MP_START Then
         
             
             cardresult = DO_WritePort(card, Channel_P1A, &H0)  ' sel chip
             Call PowerSet(500)   ' close power to disable chip
             
 
           '  Call MsecDelay(6.5)
             
            Dim TimerCounter As Integer
            Dim TmpString As String
            
             
            Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName("vid")
            Loop While TmpString = "" And TimerCounter < 150
             
             
             If TmpString = "" Then   ' can not find device after 15 s
             
               TestResult = "Bin2"
               MPTester.TestResultLab = "Bin2:MP UNKNOW Fail when enter MP"
               Exit Sub
             End If
             
             Call MsecDelay(2.5)
               
             MPTester.Print " MP Begin....."
             
             Call StartMP_Click_AU6988
   
             
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
                
               ' time out fail
                If PassTime > 65 Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    '(1)
                    winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                      Loop While winHwnd <> 0
                    End If
                    
                    '(2)
                      winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    Exit Sub
                End If
                
                ' MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    
                    winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                      Loop While winHwnd <> 0
                    End If
                    
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    
                    Exit Sub
                End If
                
                
                 'unknow fail
                 If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                     MPContFail = MPContFail + 1
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                    MPTester.Print "MP UNKNOW Fail"
                    
                   winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                     Do
                     rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                     Call MsecDelay(0.5)
                     winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                     Loop While winHwnd <> 0
                     End If
                     
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                     If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                     End If
                     
                     
                     
                    Exit Sub
                End If
                 
                
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_PASS Then
                     MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPTester.Print "MP PASS"
                End If
        End If
   
End If
'=========================================
 '    Close MP program
 '=========================================
winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
If winHwnd <> 0 Then
  Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
  Loop While winHwnd <> 0
End If
    
 winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
    Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
    Loop While winHwnd <> 0
End If

  Dim pid As Long          ' unload driver
  Dim hProcess As Long
  Dim ExitEvent As Long
 
  pid = Shell(App.Path & "\AlcorMP_698x_PD\loaddrv.exe uninstall_058F6387")
  hProcess = OpenProcess(SYNCHRONIZE + PROCESS_QUERY_INFORMATION + PROCESS_TERMINATE, 0, pid)
  ExitEvent = WaitForSingleObject(hProcess, INFINITE)
  Call CloseHandle(hProcess)
 
 
 
                        
 '=========================================
 '    POWER on
 '=========================================
 
 If MPFlag = 1 Then
        Call PowerSet(3)
          cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        
         Call MsecDelay(0.5)  ' power of to unload MPDriver

           cardresult = DO_WritePort(card, Channel_P1A, &H0)  'sel socket
          Call PowerSet(1500)
     
        
         Call MsecDelay(1.2)
        MPFlag = 0
 Else
          cardresult = DO_WritePort(card, Channel_P1A, &H0)
         Call PowerSet(1500)
         
         Call MsecDelay(1.2)
End If
         Call LoadRWTest_Click_AU6988

        
        
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
        
       If PassTime > 5 Then
           TestResult = "Bin3"
           MPTester.TestResultLab = "Bin3:RW Ready Fail"
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
       
            Exit Sub
       End If
         
         
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        Call StartRWTest_Click_AU6988
        
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
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
            
              cardresult = DO_WritePort(card, Channel_P1A, &HFB)  ' power off
            
            Call PowerSet(1500)
        
            
       
            Exit Sub
        End If
        
        
     
               
        Select Case AlcorMPMessage
        
        Case WM_FT_RW_UNKNOW_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:UnKnow Fail"
          
             ContFail = ContFail + 1
        
        Case WM_FT_RW_SPEED_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:SPEED Error "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_RW_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:RW FAIL "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_ROM_FAIL
              TestResult = "Bin4"
              MPTester.TestResultLab = "Bin4:ROM FAIL "
              ContFail = ContFail + 1
              
        Case WM_FT_RW_RAM_FAIL
              TestResult = "Bin5"
              MPTester.TestResultLab = "Bin5:RAM FAIL "
               ContFail = ContFail + 1
        Case WM_FT_RW_RW_PASS
        
               
               For LedCount = 1 To 20
               Call MsecDelay(0.1)
               cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
                If LightOn = 254 Then
               
                 Exit For
               End If
               Next LedCount
                 
                  MPTester.Print "light="; LightOn
                 If LightOn = 254 Then
                    MPTester.TestResultLab = "PASS "
                    TestResult = "PASS"
                    ContFail = 0
                Else
                 
                  TestResult = "Bin3"
                  MPTester.TestResultLab = "Bin3:LED FAIL "
              
               End If
               
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:Undefine Fail"
          
             ContFail = ContFail + 1
        
               
        End Select
                               
       cardresult = DO_WritePort(card, Channel_P1A, &HFB)
        Call PowerSet(1500)
        
         
                            
End Sub

Public Sub AU6988H56HLF2ATestSub()
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
 
If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
    FileCopy App.Path & "\AU6986AU6988_20090904\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
    Call MsecDelay(5)
End If


NewChipFlag = 0
If OldChipName <> ChipName Then
            FileCopy App.Path & "\AU6986AU6988_20090904\ROM\" & ChipName & "\ROM.Hex", App.Path & "\AU6986AU6988_20090904\ROM.Hex"
            FileCopy App.Path & "\AU6986AU6988_20090904\RAM\" & ChipName & "\RAM.Bin", App.Path & "\AU6986AU6988_20090904\RAM.Bin"
            FileCopy App.Path & "\AU6986AU6988_20090904\INI\" & ChipName & "\AlcorMP.ini", App.Path & "\AU6986AU6988_20090904\AlcorMP.ini"
            NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName
 

'==============================================================
' when begin RW Test, must clear MP program
'===============================================================
'(1)
winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
  Loop While winHwnd <> 0
End If
'(2)
winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
  Loop While winHwnd <> 0
End If
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail
 
 
 '====================================
 '  Fix Card
 '====================================
 
 If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag = True) Then
    
     If MPTester.NoMP.Value = 1 Then
        
        If (NewChipFlag = 0) And (MPTester.Check1.Value = 0) Then  ' force condition
           GoTo RW_Test_Label
        End If
    End If
       
    
    If MPTester.ResetMPFailCounter.Value = 1 Then
       ContFail = 0
    End If
    
 '==============================================================
' when begin MP, must close RW program
'===============================================================
   MPFlag = 1
 
    winHwnd = FindWindow(vbNullString, "UFD Test")
    If winHwnd <> 0 Then
      Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "UFD Test")
      Loop While winHwnd <> 0
    End If
 
       '  power on
       cardresult = DO_WritePort(card, Channel_P1A, &HFF)
       Call PowerSet(3)   ' close power to disable chip
       Call MsecDelay(0.5)  ' power for load MPDriver
       MPTester.Print "wait for MP Ready"
       Call LoadMP_Click_AU6988_20090904
 
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
        Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 30 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3
        '====================================================
        If PassTime > 30 Then
          '(1)
           MPTester.Print "MP Ready Fail"
            winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
              Loop While winHwnd <> 0
            End If
           '(2)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
            If winHwnd <> 0 Then
              Do
              rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
              Call MsecDelay(0.5)
              winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
              Loop While winHwnd <> 0
            End If
            
        
             MPTester.TestResultLab = "Bin3:MP Ready Fail"
             TestResult = "Bin3"
             MPTester.Print "MP Ready Fail"
     
              
            Exit Sub
        End If
        
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_MP_START Then
         
             
             cardresult = DO_WritePort(card, Channel_P1A, &HFB)  ' sel chip
             Call PowerSet(500)   ' close power to disable chip
             
 
           '  Call MsecDelay(6.5)
             
            Dim TimerCounter As Integer
            Dim TmpString As String
            
             
            Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName("vid")
            Loop While TmpString = "" And TimerCounter < 150
             
             
             If TmpString = "" Then   ' can not find device after 15 s
             
               TestResult = "Bin2"
               MPTester.TestResultLab = "Bin2:MP UNKNOW Fail when enter MP"
               Exit Sub
             End If
             
             Call MsecDelay(2.5)
               
             MPTester.Print " MP Begin....."
             
             Call StartMP_Click_AU6988
   
             
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
                
               ' time out fail
                If PassTime > 65 Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    '(1)
                    winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                      Loop While winHwnd <> 0
                    End If
                    
                    '(2)
                      winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    Exit Sub
                End If
                
                ' MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    
                    winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                      Loop While winHwnd <> 0
                    End If
                    
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    
                    Exit Sub
                End If
                
                
                 'unknow fail
                 If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                     MPContFail = MPContFail + 1
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                    MPTester.Print "MP UNKNOW Fail"
                    
                   winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                     Do
                     rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                     Call MsecDelay(0.5)
                     winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                     Loop While winHwnd <> 0
                     End If
                     
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                     If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                     End If
                     
                     
                     
                    Exit Sub
                End If
                 
                
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_PASS Then
                     MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPTester.Print "MP PASS"
                End If
        End If
   
End If
'=========================================
 '    Close MP program
 '=========================================
winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
If winHwnd <> 0 Then
  Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
  Loop While winHwnd <> 0
End If
    
 winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
    Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
    Loop While winHwnd <> 0
End If

  Dim pid As Long          ' unload driver
  Dim hProcess As Long
  Dim ExitEvent As Long
 
  pid = Shell(App.Path & "\AU6986AU6988_20090904\loaddrv.exe uninstall_058F6387")
  hProcess = OpenProcess(SYNCHRONIZE + PROCESS_QUERY_INFORMATION + PROCESS_TERMINATE, 0, pid)
  ExitEvent = WaitForSingleObject(hProcess, INFINITE)
  Call CloseHandle(hProcess)
 
 
 
                        
 '=========================================
 '    POWER on
 '=========================================
RW_Test_Label:
 If MPFlag = 1 Then
        Call PowerSet(3)
          cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        
         Call MsecDelay(0.5)  ' power of to unload MPDriver

           cardresult = DO_WritePort(card, Channel_P1A, &HFB)  'sel socket
          Call PowerSet(1900)
     
        
         Call MsecDelay(1.2)
        MPFlag = 0
 Else
          cardresult = DO_WritePort(card, Channel_P1A, &HFB)
         Call PowerSet(1900)
         
         Call MsecDelay(1.2)
End If
         Call LoadRWTest_Click_AU6988_20090904

        
        
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
        
       If PassTime > 5 Then
           TestResult = "Bin3"
           MPTester.TestResultLab = "Bin3:RW Ready Fail"
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
       
            Exit Sub
       End If
         
         
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        Call StartRWTest_Click_AU6988
        
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
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
            
              cardresult = DO_WritePort(card, Channel_P1A, &HFB)  ' power off
            
            Call PowerSet(1900)
        
            
       
            Exit Sub
        End If
        
        
     
               
        Select Case AlcorMPMessage
        
        Case WM_FT_RW_UNKNOW_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:UnKnow Fail"
          
             ContFail = ContFail + 1
        
        Case WM_FT_RW_SPEED_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:SPEED Error "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_RW_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:RW FAIL "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_ROM_FAIL
              TestResult = "Bin4"
              MPTester.TestResultLab = "Bin4:ROM FAIL "
              ContFail = ContFail + 1
              
        Case WM_FT_RW_RAM_FAIL
              TestResult = "Bin5"
              MPTester.TestResultLab = "Bin5:RAM FAIL "
               ContFail = ContFail + 1
        Case WM_FT_RW_RW_PASS
        
               
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
              
               End If
               
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:Undefine Fail"
          
             ContFail = ContFail + 1
        
               
        End Select
                               
       cardresult = DO_WritePort(card, Channel_P1A, &HFB)
        Call PowerSet(1900)
        
         
                            
End Sub


Public Sub AU6988H55HLF28TestSub()
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
 
If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
    FileCopy App.Path & "\AlcorMP_698x_PD\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
    Call MsecDelay(5)
End If


NewChipFlag = 0
If OldChipName <> ChipName Then
            FileCopy App.Path & "\AlcorMP_698x_PD\ROM\" & ChipName & "\ROM.Hex", App.Path & "\AlcorMP_698x_PD\ROM.Hex"
            FileCopy App.Path & "\AlcorMP_698x_PD\RAM\" & ChipName & "\RAM.Bin", App.Path & "\AlcorMP_698x_PD\RAM.Bin"
            FileCopy App.Path & "\AlcorMP_698x_PD\INI\" & ChipName & "\AlcorMP.ini", App.Path & "\AlcorMP_698x_PD\AlcorMP.ini"
            NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName
 

'==============================================================
' when begin RW Test, must clear MP program
'===============================================================
'(1)
winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
  Loop While winHwnd <> 0
End If
'(2)
winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
  Loop While winHwnd <> 0
End If
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail
 
 
 '====================================
 '  Fix Card
 '====================================
 
 If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag = True) Then
    
 '==============================================================
' when begin MP, must close RW program
'===============================================================
   MPFlag = 1
 
    winHwnd = FindWindow(vbNullString, "UFD Test")
    If winHwnd <> 0 Then
      Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "UFD Test")
      Loop While winHwnd <> 0
    End If
 
       '  power on
       cardresult = DO_WritePort(card, Channel_P1A, &HFF)
       Call PowerSet(3)   ' close power to disable chip
       Call MsecDelay(0.5)  ' power for load MPDriver
       MPTester.Print "wait for MP Ready"
       Call LoadMP_Click_AU6988
 
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
        Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 30 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3
        '====================================================
        If PassTime > 30 Then
          '(1)
           MPTester.Print "MP Ready Fail"
            winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
              Loop While winHwnd <> 0
            End If
           '(2)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
            If winHwnd <> 0 Then
              Do
              rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
              Call MsecDelay(0.5)
              winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
              Loop While winHwnd <> 0
            End If
            
        
             MPTester.TestResultLab = "Bin3:MP Ready Fail"
             TestResult = "Bin3"
             MPTester.Print "MP Ready Fail"
     
              
            Exit Sub
        End If
        
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_MP_START Then
         
             
             cardresult = DO_WritePort(card, Channel_P1A, &HFB)  ' sel chip
             Call PowerSet(500)   ' close power to disable chip
             
 
           '  Call MsecDelay(6.5)
             
            Dim TimerCounter As Integer
            Dim TmpString As String
            
             
            Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName("vid")
            Loop While TmpString = "" And TimerCounter < 150
             
             
             If TmpString = "" Then   ' can not find device after 15 s
             
               TestResult = "Bin2"
               MPTester.TestResultLab = "Bin2:MP UNKNOW Fail when enter MP"
               Exit Sub
             End If
             
             Call MsecDelay(2.5)
               
             MPTester.Print " MP Begin....."
             
             Call StartMP_Click_AU6988
   
             
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
                
               ' time out fail
                If PassTime > 65 Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    '(1)
                    winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                      Loop While winHwnd <> 0
                    End If
                    
                    '(2)
                      winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    Exit Sub
                End If
                
                ' MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    
                    winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                      Loop While winHwnd <> 0
                    End If
                    
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    
                    Exit Sub
                End If
                
                
                 'unknow fail
                 If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                     MPContFail = MPContFail + 1
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                    MPTester.Print "MP UNKNOW Fail"
                    
                   winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                     Do
                     rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                     Call MsecDelay(0.5)
                     winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                     Loop While winHwnd <> 0
                     End If
                     
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                     If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                     End If
                     
                     
                     
                    Exit Sub
                End If
                 
                
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_PASS Then
                     MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPTester.Print "MP PASS"
                End If
        End If
   
End If
'=========================================
 '    Close MP program
 '=========================================
winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
If winHwnd <> 0 Then
  Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
  Loop While winHwnd <> 0
End If
    
 winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
    Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
    Loop While winHwnd <> 0
End If

  Dim pid As Long          ' unload driver
  Dim hProcess As Long
  Dim ExitEvent As Long
 
  pid = Shell(App.Path & "\AlcorMP_698x_PD\loaddrv.exe uninstall_058F6387")
  hProcess = OpenProcess(SYNCHRONIZE + PROCESS_QUERY_INFORMATION + PROCESS_TERMINATE, 0, pid)
  ExitEvent = WaitForSingleObject(hProcess, INFINITE)
  Call CloseHandle(hProcess)
 
 
 
                        
 '=========================================
 '    POWER on
 '=========================================
 
 If MPFlag = 1 Then
        Call PowerSet(3)
          cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        
         Call MsecDelay(0.5)  ' power of to unload MPDriver

           cardresult = DO_WritePort(card, Channel_P1A, &HFB)  'sel socket
          Call PowerSet(1900)
     
        
         Call MsecDelay(1.2)
        MPFlag = 0
 Else
          cardresult = DO_WritePort(card, Channel_P1A, &HFB)
         Call PowerSet(1900)
         
         Call MsecDelay(1.2)
End If
         Call LoadRWTest_Click_AU6988

        
        
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
        
       If PassTime > 5 Then
           TestResult = "Bin3"
           MPTester.TestResultLab = "Bin3:RW Ready Fail"
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
       
            Exit Sub
       End If
         
         
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        Call StartRWTest_Click_AU6988
        
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
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
            
              cardresult = DO_WritePort(card, Channel_P1A, &HFB)  ' power off
            
            Call PowerSet(1900)
        
            
       
            Exit Sub
        End If
        
        
     
               
        Select Case AlcorMPMessage
        
        Case WM_FT_RW_UNKNOW_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:UnKnow Fail"
          
             ContFail = ContFail + 1
        
        Case WM_FT_RW_SPEED_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:SPEED Error "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_RW_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:RW FAIL "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_ROM_FAIL
              TestResult = "Bin4"
              MPTester.TestResultLab = "Bin4:ROM FAIL "
              ContFail = ContFail + 1
              
        Case WM_FT_RW_RAM_FAIL
              TestResult = "Bin5"
              MPTester.TestResultLab = "Bin5:RAM FAIL "
               ContFail = ContFail + 1
        Case WM_FT_RW_RW_PASS
        
               
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
              
               End If
               
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:Undefine Fail"
          
             ContFail = ContFail + 1
        
               
        End Select
                               
       cardresult = DO_WritePort(card, Channel_P1A, &HFB)
        Call PowerSet(1900)
        
         
                            
End Sub

Public Sub AU6988D52HLF26TestSub()
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
 
If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
    FileCopy App.Path & "\AlcorMP_698x_PD\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
    Call MsecDelay(5)
End If


NewChipFlag = 0
If OldChipName <> ChipName Then
            FileCopy App.Path & "\AlcorMP_698x_PD\ROM\" & ChipName & "\ROM.Hex", App.Path & "\AlcorMP_698x_PD\ROM.Hex"
            FileCopy App.Path & "\AlcorMP_698x_PD\RAM\" & ChipName & "\RAM.Bin", App.Path & "\AlcorMP_698x_PD\RAM.Bin"
            FileCopy App.Path & "\AlcorMP_698x_PD\INI\" & ChipName & "\AlcorMP.ini", App.Path & "\AlcorMP_698x_PD\AlcorMP.ini"
            NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName
 

'==============================================================
' when begin RW Test, must clear MP program
'===============================================================
'(1)
winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
  Loop While winHwnd <> 0
End If
'(2)
winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
  Loop While winHwnd <> 0
End If
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail
 
 
 '====================================
 '  Fix Card
 '====================================
 
 If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag = True) Then
    
 '==============================================================
' when begin MP, must close RW program
'===============================================================
   MPFlag = 1
 
    winHwnd = FindWindow(vbNullString, "UFD Test")
    If winHwnd <> 0 Then
      Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "UFD Test")
      Loop While winHwnd <> 0
    End If
 
       '  power on
       cardresult = DO_WritePort(card, Channel_P1A, &HFF)
       Call PowerSet(3)   ' close power to disable chip
       Call MsecDelay(0.5)  ' power for load MPDriver
       MPTester.Print "wait for MP Ready"
       Call LoadMP_Click_AU6988
 
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
        Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 30 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3
        '====================================================
        If PassTime > 30 Then
          '(1)
           MPTester.Print "MP Ready Fail"
            winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
              Loop While winHwnd <> 0
            End If
           '(2)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
            If winHwnd <> 0 Then
              Do
              rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
              Call MsecDelay(0.5)
              winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
              Loop While winHwnd <> 0
            End If
            
        
             MPTester.TestResultLab = "Bin3:MP Ready Fail"
             TestResult = "Bin3"
             MPTester.Print "MP Ready Fail"
     
              
            Exit Sub
        End If
        
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_MP_START Then
         
             
             cardresult = DO_WritePort(card, Channel_P1A, &HFB)  ' sel chip
             Call PowerSet(500)   ' close power to disable chip
             
 
           '  Call MsecDelay(6.5)
             
            Dim TimerCounter As Integer
            Dim TmpString As String
            
             
            Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName("vid")
            Loop While TmpString = "" And TimerCounter < 150
             
             
             If TmpString = "" Then   ' can not find device after 15 s
             
               TestResult = "Bin2"
               MPTester.TestResultLab = "Bin2:MP UNKNOW Fail when enter MP"
               Exit Sub
             End If
             
             Call MsecDelay(2.5)
               
             MPTester.Print " MP Begin....."
             
             Call StartMP_Click_AU6988
   
             
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
                
               ' time out fail
                If PassTime > 65 Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    '(1)
                    winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                      Loop While winHwnd <> 0
                    End If
                    
                    '(2)
                      winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    Exit Sub
                End If
                
                ' MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    
                    winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                      Loop While winHwnd <> 0
                    End If
                    
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    
                    Exit Sub
                End If
                
                
                 'unknow fail
                 If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                     MPContFail = MPContFail + 1
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                    MPTester.Print "MP UNKNOW Fail"
                    
                   winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                    If winHwnd <> 0 Then
                     Do
                     rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                     Call MsecDelay(0.5)
                     winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
                     Loop While winHwnd <> 0
                     End If
                     
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                     If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                     End If
                     
                     
                     
                    Exit Sub
                End If
                 
                
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_PASS Then
                     MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPTester.Print "MP PASS"
                End If
        End If
   
End If
'=========================================
 '    Close MP program
 '=========================================
winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
If winHwnd <> 0 Then
  Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
  Loop While winHwnd <> 0
End If
    
 winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
    Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
    Loop While winHwnd <> 0
End If

  Dim pid As Long          ' unload driver
  Dim hProcess As Long
  Dim ExitEvent As Long
 
  pid = Shell(App.Path & "\AlcorMP_698x_PD\loaddrv.exe uninstall_058F6387")
  hProcess = OpenProcess(SYNCHRONIZE + PROCESS_QUERY_INFORMATION + PROCESS_TERMINATE, 0, pid)
  ExitEvent = WaitForSingleObject(hProcess, INFINITE)
  Call CloseHandle(hProcess)
 
 
 
                        
 '=========================================
 '    POWER on
 '=========================================
 
 If MPFlag = 1 Then
        Call PowerSet(3)
          cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        
         Call MsecDelay(0.5)  ' power of to unload MPDriver

           cardresult = DO_WritePort(card, Channel_P1A, &HFB)  'sel socket
          Call PowerSet(1500)
     
        
         Call MsecDelay(1.2)
        MPFlag = 0
 Else
          cardresult = DO_WritePort(card, Channel_P1A, &HFB)
         Call PowerSet(1500)
         
         Call MsecDelay(1.2)
End If
         Call LoadRWTest_Click_AU6988

        
        
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
        
       If PassTime > 5 Then
           TestResult = "Bin3"
           MPTester.TestResultLab = "Bin3:RW Ready Fail"
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
       
            Exit Sub
       End If
         
         
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        Call StartRWTest_Click_AU6988
        
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
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
            
              cardresult = DO_WritePort(card, Channel_P1A, &HFB)  ' power off
            
            Call PowerSet(1500)
        
            
       
            Exit Sub
        End If
        
        
     
               
        Select Case AlcorMPMessage
        
        Case WM_FT_RW_UNKNOW_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:UnKnow Fail"
          
             ContFail = ContFail + 1
        
        Case WM_FT_RW_SPEED_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:SPEED Error "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_RW_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:RW FAIL "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_ROM_FAIL
              TestResult = "Bin4"
              MPTester.TestResultLab = "Bin4:ROM FAIL "
              ContFail = ContFail + 1
              
        Case WM_FT_RW_RAM_FAIL
              TestResult = "Bin5"
              MPTester.TestResultLab = "Bin5:RAM FAIL "
               ContFail = ContFail + 1
        Case WM_FT_RW_RW_PASS
        
               
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
              
               End If
               
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:Undefine Fail"
          
             ContFail = ContFail + 1
        
               
        End Select
                               
       cardresult = DO_WritePort(card, Channel_P1A, &HFB)
        Call PowerSet(1500)
        
         
                            
End Sub

Public Sub AU6992A53DLF21TestSub()

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
 Dim TmpChipName
 
 TmpChipName = ChipName
 ChipName = "AU6992A53DLF20"
 
   MPTester.TestResultLab = ""
'===============================================================
' Fail location initial
'===============================================================
 
If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
    FileCopy App.Path & "\AlcorMP_698x2\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
    Call MsecDelay(5)
End If


NewChipFlag = 0
If OldChipName <> ChipName Then
            FileCopy App.Path & "\AlcorMP_698x2\ROM\" & ChipName & "\ROM.Hex", App.Path & "\AlcorMP_698x2\ROM.Hex"
            FileCopy App.Path & "\AlcorMP_698x2\RAM\" & ChipName & "\RAM.Bin", App.Path & "\AlcorMP_698x2\RAM.Bin"
            FileCopy App.Path & "\AlcorMP_698x2\INI\" & ChipName & "\AlcorMP.ini", App.Path & "\AlcorMP_698x2\AlcorMP.ini"
            NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName
 

'==============================================================
' when begin RW Test, must clear MP program
'===============================================================


'(1)  /// for Auto mode
winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
  Loop While winHwnd <> 0
End If

'(1)
winHwnd = FindWindow(vbNullString, AU6992MPCaption)
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, AU6992MPCaption)
  Loop While winHwnd <> 0
End If
'(2)
winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
  Loop While winHwnd <> 0
End If
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail
 
 
 '====================================
 '  Fix Card
 '====================================
' GoTo T1
 If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag = True) Then
 
 
   If MPTester.NoMP.Value = 1 Then
        
        If (NewChipFlag = 0) And (MPTester.Check1.Value = 0) Then  ' force condition
           GoTo RW_Test_Label
        End If
    End If
       
    
    If MPTester.ResetMPFailCounter.Value = 1 Then
       ContFail = 0
    End If
    
 '==============================================================
' when begin MP, must close RW program
'===============================================================
   MPFlag = 1
 
    winHwnd = FindWindow(vbNullString, "UFD Test")
    If winHwnd <> 0 Then
      Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "UFD Test")
      Loop While winHwnd <> 0
    End If
 
       '  power on
       cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        Call PowerSet(3)   ' close power to disable chip
       Call MsecDelay(0.5)  ' power for load MPDriver
       MPTester.Print "wait for MP Ready"
       Call LoadMP_Click_AU6992
 
        OldTimer = Timer
        AlcorMPMessage = 0
        Debug.Print "begin"
        Do
           ' DoEvents
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
            PassTime = Timer - OldTimer
            'Debug.Print AlcorMPMessage
        Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 30 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3
        '====================================================
        If PassTime > 30 Then
          '(1)
           MPTester.Print "MP Ready Fail"
            winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
              Loop While winHwnd <> 0
            End If
           '(2)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
            If winHwnd <> 0 Then
              Do
              rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
              Call MsecDelay(0.5)
              winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
              Loop While winHwnd <> 0
            End If
            
        
             MPTester.TestResultLab = "Bin3:MP Ready Fail"
             TestResult = "Bin3"
             MPTester.Print "MP Ready Fail"
     
              
            Exit Sub
        End If
        
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_MP_START Then
         
             
             cardresult = DO_WritePort(card, Channel_P1A, &HFD)  ' sel chip
              Call PowerSet(500)   ' close power to disable chip
             
 
          
             
            Dim TimerCounter As Integer
            Dim TmpString As String
            
             
            Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName("vid")
            Loop While TmpString = "" And TimerCounter < 150
             
             
             If TmpString = "" Then   ' can not find device after 15 s
             
               TestResult = "Bin2"
               MPTester.TestResultLab = "Bin2:MP UNKNOW Fail when enter MP"
               Exit Sub
             End If
             
             Call MsecDelay(2.5)
               
             MPTester.Print " MP Begin....."
             
             Call StartMP_Click_AU6992
   
             
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
                
               ' time out fail
                If PassTime > 65 Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    '(1)
                    winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                      Loop While winHwnd <> 0
                    End If
                    
                    '(2)
                      winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    Exit Sub
                End If
                
                ' MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    
                    winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                      Loop While winHwnd <> 0
                    End If
                    
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    
                    Exit Sub
                End If
                
                
                 'unknow fail
                 If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                     MPContFail = MPContFail + 1
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                    MPTester.Print "MP UNKNOW Fail"
                    
                   winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                    If winHwnd <> 0 Then
                     Do
                     rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                     Call MsecDelay(0.5)
                     winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                     Loop While winHwnd <> 0
                     End If
                     
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                     If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                     End If
                     
                     
                     
                    Exit Sub
                End If
                 
                
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_PASS Then
                     MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPTester.Print "MP PASS"
                End If
        End If
   
End If
'=========================================
 '    Close MP program
 '=========================================
winHwnd = FindWindow(vbNullString, AU6992MPCaption)
If winHwnd <> 0 Then
  Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, AU6992MPCaption)
  Loop While winHwnd <> 0
End If
    
 winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
    Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
    Loop While winHwnd <> 0
End If

  Dim pid As Long          ' unload driver
  Dim hProcess As Long
  Dim ExitEvent As Long
 
  pid = Shell(App.Path & "\AlcorMP_698x2\loaddrv.exe uninstall_058F6387")
  hProcess = OpenProcess(SYNCHRONIZE + PROCESS_QUERY_INFORMATION + PROCESS_TERMINATE, 0, pid)
  ExitEvent = WaitForSingleObject(hProcess, INFINITE)
  Call CloseHandle(hProcess)
  KillProcess ("LoadDrv.exe")
 
 
                        
 '=========================================
 '    POWER on
 '=========================================
'T1:
RW_Test_Label:
 
 If MPFlag = 1 Then
         Call PowerSet(3)
          cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        
         Call MsecDelay(0.5)  ' power of to unload MPDriver

           cardresult = DO_WritePort(card, Channel_P1A, &HFA)  'sel socket
           Call PowerSet(1500)
     
        
         Call MsecDelay(1.2)
        MPFlag = 0
 Else
          cardresult = DO_WritePort(card, Channel_P1A, &HFA)
         Call PowerSet(1500)
         
         Call MsecDelay(1.2)
End If
         Call LoadRWTest_Click_AU6992

        
        
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
     '   GoTo T2
       If PassTime > 5 Then
           TestResult = "Bin3"
           MPTester.TestResultLab = "Bin3:RW Ready Fail"
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
       
            Exit Sub
       End If
         
T2:
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        Call StartRWTest_Click_AU6988
        
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
                Or AlcorMPMessage = WM_FT_CHECK_CERBGPO_FAIL _
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
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
            
              cardresult = DO_WritePort(card, Channel_P1A, &HFA)  ' power off
            
             Call PowerSet(1500)
        
            
       
            Exit Sub
        End If
        
        
     
               
        Select Case AlcorMPMessage
        
        Case WM_FT_RW_UNKNOW_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:UnKnow Fail"
          
             ContFail = ContFail + 1
        
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
             
        Case WM_FT_RW_ROM_FAIL
              TestResult = "Bin4"
              MPTester.TestResultLab = "Bin4:ROM FAIL "
              ContFail = ContFail + 1
              
        Case WM_FT_RW_RAM_FAIL
              TestResult = "Bin5"
              MPTester.TestResultLab = "Bin5:RAM FAIL "
               ContFail = ContFail + 1
        Case WM_FT_RW_RW_PASS
        
               
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
              
               End If
               
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:Undefine Fail"
          
             ContFail = ContFail + 1
        
               
        End Select
                               
       cardresult = DO_WritePort(card, Channel_P1A, &HFA)
         Call PowerSet(1500)
        
         
                            
                            
End Sub

Public Sub AU6992_MP_Golden()

    If ChipName = "AU6992A53DLF2B" Then
        ChipName = "AU6992A53DLF2A"
    End If
    
    Call AU6992DualTestSub
    
End Sub

Public Sub AU6992EQCReNameSub()

    If ChipName = "AU6992A53HLF03" Then
        ChipName = "AU6992A53HLF23"
    ElseIf ChipName = "AU6992B53HLF03" Then
        ChipName = "AU6992B53HLF23"
    ElseIf ChipName = "AU6992R53HLF03" Then
        ChipName = "AU6992R53HLF23"
    ElseIf ChipName = "AU6992S53HLF03" Then
        ChipName = "AU6992S53HLF23"
    Else
        ChipName = ChipName
    End If
    
    Call AU6992EQCHWSingleTestSub
    
End Sub

Public Sub AU6992EQCReNameDualSub()

    If ChipName = "AU6992A53HLF0C" Then
        ChipName = "AU6992A53HLF23"
    ElseIf ChipName = "AU6992B53HLF0C" Then
        ChipName = "AU6992B53HLF23"
    ElseIf ChipName = "AU6992R53HLF0C" Then
        ChipName = "AU6992R53HLF23"
    ElseIf ChipName = "AU6992S53HLF0C" Then
        ChipName = "AU6992S53HLF23"
    Else
        ChipName = ChipName
    End If
    
    Call AU6992EQCHW_DualTestSub
    
End Sub

Public Sub AU6992DualTestSub()

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
 
MPTester.TestResultLab = ""
'===============================================================
' Fail location initial
'===============================================================
 
If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
    FileCopy App.Path & "\AlcorMP_698x2\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
    Call MsecDelay(5)
End If

NewChipFlag = 0
If OldChipName <> ChipName Then
    FileCopy App.Path & "\AlcorMP_698x2\New_INI\" & ChipName & "\ROM\ROM.Hex", App.Path & "\AlcorMP_698x2\ROM.Hex"
    FileCopy App.Path & "\AlcorMP_698x2\New_INI\" & ChipName & "\RAM\\RAM.Bin", App.Path & "\AlcorMP_698x2\RAM.Bin"
    FileCopy App.Path & "\AlcorMP_698x2\New_INI\" & ChipName & "\AlcorMP.ini", App.Path & "\AlcorMP_698x2\AlcorMP.ini"
    FileCopy App.Path & "\AlcorMP_698x2\New_INI\" & ChipName & "\PE.bin", App.Path & "\AlcorMP_698x2\PE.bin"
    NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName
 

'==============================================================
' when begin RW Test, must clear MP program
'===============================================================


'(1)  /// for Auto mode
winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
  Loop While winHwnd <> 0
End If

'(1)
winHwnd = FindWindow(vbNullString, AU6992MPCaption)
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, AU6992MPCaption)
  Loop While winHwnd <> 0
End If
'(2)
winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
  Loop While winHwnd <> 0
End If
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail
 
 
 '====================================
 '  Fix Card
 '====================================
' GoTo T1
 If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag = True) Then
 
 
   If MPTester.NoMP.Value = 1 Then
        
        If (NewChipFlag = 0) And (MPTester.Check1.Value = 0) Then  ' force condition
           GoTo RW_Test_Label
        End If
    End If
       
    
    If MPTester.ResetMPFailCounter.Value = 1 Then
       ContFail = 0
    End If
    
 '==============================================================
' when begin MP, must close RW program
'===============================================================
   MPFlag = 1
 
    winHwnd = FindWindow(vbNullString, "UFD Test")
    If winHwnd <> 0 Then
      Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "UFD Test")
      Loop While winHwnd <> 0
    End If
 
       '  power on
       cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        Call PowerSet(3)   ' close power to disable chip
       Call MsecDelay(0.5)  ' power for load MPDriver
       MPTester.Print "wait for MP Ready"
       Call LoadMP_Click_AU6992
 
        OldTimer = Timer
        AlcorMPMessage = 0
        Debug.Print "begin"
        Do
           ' DoEvents
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
            PassTime = Timer - OldTimer
            'Debug.Print AlcorMPMessage
        Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 30 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3
        '====================================================
        If PassTime > 30 Then
          '(1)
           MPTester.Print "MP Ready Fail"
            winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
              Loop While winHwnd <> 0
            End If
           '(2)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
            If winHwnd <> 0 Then
              Do
              rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
              Call MsecDelay(0.5)
              winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
              Loop While winHwnd <> 0
            End If
            
        
             MPTester.TestResultLab = "Bin3:MP Ready Fail"
             TestResult = "Bin3"
             MPTester.Print "MP Ready Fail"
     
              
            Exit Sub
        End If
        
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_MP_START Then
         
             
             cardresult = DO_WritePort(card, Channel_P1A, &HFD)  ' sel chip
              Call PowerSet(500)   ' close power to disable chip
             
             
            Dim TimerCounter As Integer
            Dim TmpString As String
            
             
            Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName("vid")
            Loop While TmpString = "" And TimerCounter < 150
             
            Call MsecDelay(0.3)
             
             If TmpString = "" Then   ' can not find device after 15 s
             
               TestResult = "Bin2"
               MPTester.TestResultLab = "Bin2:MP UNKNOW Fail when enter MP"
               Exit Sub
             End If
             
             Call MsecDelay(2.5)
               
             MPTester.Print " MP Begin....."
             
             Call StartMP_Click_AU6992
   
              ReMP_Flag = 0
              OldTimer = Timer
              AlcorMPMessage = 0
                
                Do
                   'DoEvents
                    If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                        AlcorMPMessage = mMsg.message
                        TranslateMessage mMsg
                        DispatchMessage mMsg
                            
                        If (AlcorMPMessage = WM_FT_MP_FAIL) And (ReMP_Flag = 0) Then
                            'ReMP_Flag = 1
                            AlcorMPMessage = 1
                            cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'close power
                            Call MsecDelay(0.3)
                            cardresult = DO_WritePort(card, Channel_P1A, &HFD)  'open power
                            Call MsecDelay(2.2)
                            Call RefreshMP_Click_AU6992
                            Call MsecDelay(0.5)
                            Call StartMP_Click_AU6992
                            
                            ReMP_Counter = ReMP_Counter + 1
                            If ReMP_Counter >= ReMP_Limit Then
                                ReMP_Flag = 1
                                ReMP_Counter = 0
                            End If
                        End If
                    End If
                    
                    PassTime = Timer - OldTimer
                
                Loop Until AlcorMPMessage = WM_FT_MP_PASS _
                Or AlcorMPMessage = WM_FT_MP_FAIL _
                Or AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL _
                Or PassTime > 65 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                
                If AlcorMPMessage = WM_FT_MP_PASS Then
                    ReMP_Counter = 0
                End If

                MPTester.Print "MP work time="; PassTime
                 MPTester.MPText.Text = Hex(AlcorMPMessage)
                '================================================
                '  Handle MP work time out error
                '===============================================
                
               ' time out fail
                If PassTime > 65 Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    '(1)
                    winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                      Loop While winHwnd <> 0
                    End If
                    
                    '(2)
                      winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    Exit Sub
                End If
                
                ' MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    
                    winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                      Loop While winHwnd <> 0
                    End If
                    
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    
                    Exit Sub
                End If
                
                
                 'unknow fail
                 If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                     MPContFail = MPContFail + 1
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                    MPTester.Print "MP UNKNOW Fail"
                    
                   winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                    If winHwnd <> 0 Then
                     Do
                     rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                     Call MsecDelay(0.5)
                     winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                     Loop While winHwnd <> 0
                     End If
                     
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                     If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                     End If
                     
                     
                     
                    Exit Sub
                End If
                 
                
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_PASS Then
                     MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPTester.Print "MP PASS"
                End If
        End If
   
End If
'=========================================
 '    Close MP program
 '=========================================
winHwnd = FindWindow(vbNullString, AU6992MPCaption)
If winHwnd <> 0 Then
  Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, AU6992MPCaption)
  Loop While winHwnd <> 0
    
    Call MsecDelay(0.2)
    KillProcess ("AlcorMP.exe")

End If
    
 winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
    Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
    Loop While winHwnd <> 0
End If

  Dim pid As Long          ' unload driver
  Dim hProcess As Long
  Dim ExitEvent As Long
 
  pid = Shell(App.Path & "\AlcorMP_698x2\loaddrv.exe uninstall_058F6387")
  hProcess = OpenProcess(SYNCHRONIZE + PROCESS_QUERY_INFORMATION + PROCESS_TERMINATE, 0, pid)
  ExitEvent = WaitForSingleObject(hProcess, INFINITE)
  Call CloseHandle(hProcess)
  KillProcess ("LoadDrv.exe")
 
 
                        
 '=========================================
 '    POWER on
 '=========================================
'T1:
RW_Test_Label:
 
 If MPFlag = 1 Then
         Call PowerSet(3)
          cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        
         Call MsecDelay(0.5)  ' power of to unload MPDriver

           cardresult = DO_WritePort(card, Channel_P1A, &HFA)  'sel socket
           Call PowerSet(1500)
     
        
         Call MsecDelay(1.2)
        MPFlag = 0
 Else
          cardresult = DO_WritePort(card, Channel_P1A, &HFA)
         Call PowerSet(1500)
         
         Call MsecDelay(1.2)
End If
         Call LoadRWTest_Click_AU6992

        
        
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
     '   GoTo T2
       If PassTime > 5 Then
           TestResult = "Bin3"
           MPTester.TestResultLab = "Bin3:RW Ready Fail"
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
       
            Exit Sub
       End If
         
T2:
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        Call StartRWTest_Click_AU6988
        
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
                Or AlcorMPMessage = WM_FT_CHECK_CERBGPO_FAIL _
              Or PassTime > 5 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
    
          MPTester.Print "RW work Time="; PassTime
          MPTester.MPText.Text = Hex(AlcorMPMessage)
        
        
        '===========================================================
        '  RW Time Out Fail
        '===========================================================
        
        If (PassTime > 5) Or ((FailCloseAP) And (AlcorMPMessage <> WM_FT_RW_RW_PASS)) Then
            Close_FT_AP ("UFD Test")
            
            If (PassTime > 5) Then
                TestResult = "Bin3"
                MPTester.TestResultLab = "Bin3:RW Time Out Fail"
                cardresult = DO_WritePort(card, Channel_P1A, &HFA)  ' power off
                Exit Sub
            End If
        
        End If
        
        
     
               
        Select Case AlcorMPMessage
        
        Case WM_FT_RW_UNKNOW_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:UnKnow Fail"
          
             ContFail = ContFail + 1
        
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
             
        Case WM_FT_RW_ROM_FAIL
              TestResult = "Bin4"
              MPTester.TestResultLab = "Bin4:ROM FAIL "
              ContFail = ContFail + 1
              
        Case WM_FT_RW_RAM_FAIL
              TestResult = "Bin5"
              MPTester.TestResultLab = "Bin5:RAM FAIL "
               ContFail = ContFail + 1
               
        Case WM_FT_RW_RW_PASS
        
               
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
              
               End If
               
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:Undefine Fail"
          
             ContFail = ContFail + 1
        
               
        End Select
                               
       cardresult = DO_WritePort(card, Channel_P1A, &HFA)
         Call PowerSet(1500)
        
         
                            
                            
End Sub

Public Sub AU6992EQCHW_DualTestSub()

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
 'Dim ReMP_Flag As Byte
 Dim HV_Result As String
 Dim LV_Result As String
 
 
    MPTester.TestResultLab = ""
    HV_Result = ""
    LV_Result = ""
    EQC_HV = False
    EQC_LV = False

 
   MPTester.TestResultLab = ""
'===============================================================
' Fail location initial
'===============================================================
 
If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
    FileCopy App.Path & "\AlcorMP_698x2\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
    Call MsecDelay(5)
End If


NewChipFlag = 0
If OldChipName <> ChipName Then
    FileCopy App.Path & "\AlcorMP_698x2\New_INI\" & ChipName & "\ROM\ROM.Hex", App.Path & "\AlcorMP_698x2\ROM.Hex"
    FileCopy App.Path & "\AlcorMP_698x2\New_INI\" & ChipName & "\RAM\\RAM.Bin", App.Path & "\AlcorMP_698x2\RAM.Bin"
    FileCopy App.Path & "\AlcorMP_698x2\New_INI\" & ChipName & "\AlcorMP.ini", App.Path & "\AlcorMP_698x2\AlcorMP.ini"
    FileCopy App.Path & "\AlcorMP_698x2\New_INI\" & ChipName & "\PE.bin", App.Path & "\AlcorMP_698x2\PE.bin"
    FileCopy App.Path & "\AlcorMP_698x2\New_INI\" & ChipName & "\FT.ini", App.Path & "\FT.ini"
    FileCopy App.Path & "\AlcorMP_698x2\New_INI\" & ChipName & "\FT.ini", App.Path & "\AlcorMP_698x2\FT.ini"
    NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName
 

'==============================================================
' when begin RW Test, must clear MP program
'===============================================================


'(1)  /// for Auto mode
winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
  Loop While winHwnd <> 0
End If

'(1)
winHwnd = FindWindow(vbNullString, AU6992MPCaption)
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, AU6992MPCaption)
  Loop While winHwnd <> 0
End If
'(2)
winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
  Loop While winHwnd <> 0
End If
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail
 
 
 '====================================
 '  Fix Card
 '====================================
' GoTo T1
 If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag = True) Then
 
 
   If MPTester.NoMP.Value = 1 Then
        
        If (NewChipFlag = 0) And (MPTester.Check1.Value = 0) Then  ' force condition
           GoTo RW_Test_Label
        End If
    End If
       
    
    If MPTester.ResetMPFailCounter.Value = 1 Then
       ContFail = 0
    End If
    
 '==============================================================
' when begin MP, must close RW program
'===============================================================
   MPFlag = 1
 
    winHwnd = FindWindow(vbNullString, "UFD Test")
    If winHwnd <> 0 Then
      Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "UFD Test")
      Loop While winHwnd <> 0
    End If
 
       '  power on
       cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        Call PowerSet(3)   ' close power to disable chip
       Call MsecDelay(0.5)  ' power for load MPDriver
       MPTester.Print "wait for MP Ready"
       Call LoadMP_Click_AU6992
 
        OldTimer = Timer
        AlcorMPMessage = 0
        Debug.Print "begin"
        Do
           ' DoEvents
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
            PassTime = Timer - OldTimer
            'Debug.Print AlcorMPMessage
        Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 30 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3
        '====================================================
        If PassTime > 30 Then
          '(1)
           MPTester.Print "MP Ready Fail"
            winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
              Loop While winHwnd <> 0
            End If
           '(2)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
            If winHwnd <> 0 Then
              Do
              rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
              Call MsecDelay(0.5)
              winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
              Loop While winHwnd <> 0
            End If
            
        
             MPTester.TestResultLab = "Bin3:MP Ready Fail"
             TestResult = "Bin3"
             MPTester.Print "MP Ready Fail"
     
              
            Exit Sub
        End If
        
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_MP_START Then
         
             
            cardresult = DO_WritePort(card, Channel_P1A, &HFD)  ' sel chip
            'Call PowerSet(500)   ' close power to disable chip
            Call PowerSet2(1, "3.3", "0.5", 1, "3.3", "0.5", 1)
            
             
            Dim TimerCounter As Integer
            Dim TmpString As String
            
             
            Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName("vid")
            Loop While TmpString = "" And TimerCounter < 150
             
            Call MsecDelay(0.3)
             
             If TmpString = "" Then   ' can not find device after 15 s
             
               TestResult = "Bin2"
               MPTester.TestResultLab = "Bin2:MP UNKNOW Fail when enter MP"
               Exit Sub
             End If
             
             Call MsecDelay(2.5)
               
             MPTester.Print " MP Begin....."
             
             Call StartMP_Click_AU6992
   
              ReMP_Flag = 0
              OldTimer = Timer
              AlcorMPMessage = 0
                
                Do
                   'DoEvents
                    If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                        AlcorMPMessage = mMsg.message
                        TranslateMessage mMsg
                        DispatchMessage mMsg
                            
                        If (AlcorMPMessage = WM_FT_MP_FAIL) And (ReMP_Flag = 0) Then
                            'ReMP_Flag = 1
                            AlcorMPMessage = 1
                            cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'close power
                            Call MsecDelay(0.3)
                            cardresult = DO_WritePort(card, Channel_P1A, &HFD)  'open power
                            Call MsecDelay(2.2)
                            Call RefreshMP_Click_AU6992
                            Call MsecDelay(0.5)
                            Call StartMP_Click_AU6992
                            
                            ReMP_Counter = ReMP_Counter + 1
                            If ReMP_Counter >= ReMP_Limit Then
                                ReMP_Flag = 1
                                ReMP_Counter = 0
                            End If
                        End If
                    End If
                    
                    PassTime = Timer - OldTimer
                
                Loop Until AlcorMPMessage = WM_FT_MP_PASS _
                Or AlcorMPMessage = WM_FT_MP_FAIL _
                Or AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL _
                Or PassTime > 65 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                
                If AlcorMPMessage = WM_FT_MP_PASS Then
                    If ReMP_Flag = 0 Then
                        MsecDelay (MPIdleTime * (ReMP_Limit - ReMP_Counter))
                    End If
                    ReMP_Counter = 0
                End If

                MPTester.Print "MP work time="; PassTime
                 MPTester.MPText.Text = Hex(AlcorMPMessage)
                '================================================
                '  Handle MP work time out error
                '===============================================
                
               ' time out fail
                If PassTime > 65 Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    '(1)
                    winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                      Loop While winHwnd <> 0
                    End If
                    
                    '(2)
                      winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    Exit Sub
                End If
                
                ' MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    
                    winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                      Loop While winHwnd <> 0
                    End If
                    
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    
                    Exit Sub
                End If
                
                
                 'unknow fail
                 If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                     MPContFail = MPContFail + 1
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                    MPTester.Print "MP UNKNOW Fail"
                    
                   winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                    If winHwnd <> 0 Then
                     Do
                     rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                     Call MsecDelay(0.5)
                     winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                     Loop While winHwnd <> 0
                     End If
                     
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                     If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                     End If
                     
                     
                     
                    Exit Sub
                End If
                 
                
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_PASS Then
                     MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPTester.Print "MP PASS"
                End If
        End If
   
End If
'=========================================
 '    Close MP program
 '=========================================
winHwnd = FindWindow(vbNullString, AU6992MPCaption)
If winHwnd <> 0 Then
  Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, AU6992MPCaption)
  Loop While winHwnd <> 0
    
    Call MsecDelay(0.2)
    KillProcess ("AlcorMP.exe")

End If
    
 winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
    Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
    Loop While winHwnd <> 0
End If

  Dim pid As Long          ' unload driver
  Dim hProcess As Long
  Dim ExitEvent As Long
 
  pid = Shell(App.Path & "\AlcorMP_698x2\loaddrv.exe uninstall_058F6387")
  hProcess = OpenProcess(SYNCHRONIZE + PROCESS_QUERY_INFORMATION + PROCESS_TERMINATE, 0, pid)
  ExitEvent = WaitForSingleObject(hProcess, INFINITE)
  Call CloseHandle(hProcess)
  KillProcess ("LoadDrv.exe")
 
 
                        
 '=========================================
 '    POWER on
 '=========================================
'T1:
RW_Test_Label:
 
 If MPFlag = 1 Then
        
    If (EQC_HV = False) And (EQC_LV = False) Then
        Call PowerSet2(1, "3.6", "0.15", 1, "3.6", "0.15", 1)
        cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        Call MsecDelay(0.2)  ' power of to unload MPDriver

        cardresult = DO_WritePort(card, Channel_P1A, &HFA)  'sel socket
        Call MsecDelay(1.3)
        'WaitDevOn ("vid_058f")
        'Call MsecDelay(0.2)
        EQC_HV = True
    End If
     
    MPFlag = 0
 
Else
          
    If (EQC_HV = False) And (EQC_LV = False) Then
        Call PowerSet2(1, "3.6", "0.15", 1, "3.6", "0.15", 1)
        cardresult = DO_WritePort(card, Channel_P1A, &HFA)
        Call MsecDelay(1.3)
        'WaitDevOn ("vid_058f")
        'Call MsecDelay(0.2)
        EQC_HV = True
    End If

End If
        
        Call LoadRWTest_Click_AU6992

        
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
     '   GoTo T2
       If PassTime > 5 Then
           TestResult = "Bin3"
           MPTester.TestResultLab = "Bin3:RW Ready Fail"
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
       
            Exit Sub
       End If
         
T2:
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        Call StartRWTest_Click_AU6988
        
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
            Or AlcorMPMessage = WM_FT_CHECK_CERBGPO_FAIL _
            Or AlcorMPMessage = WM_FT_CHECK_HW_CODE_FAIL _
            Or AlcorMPMessage = WM_FT_PHYREAD_FAIL _
            Or AlcorMPMessage = WM_FT_ECC_FAIL _
            Or AlcorMPMessage = WM_FT_NOFREEBLOCK_FAIL _
            Or AlcorMPMessage = WM_FT_LODECODE_FAIL _
            Or AlcorMPMessage = WM_FT_RELOADCODE_FAIL _
            Or AlcorMPMessage = WM_FT_TESTUNITREADY_FAIL _
            Or PassTime > 10 _
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
            MPTester.TestResultLab = "HV: UnKnow Fail"
            'ContFail = ContFail + 1
        
        Case WM_FT_CHECK_HW_CODE_FAIL
             TestResult = "Bin5"
             MPTester.TestResultLab = "HV: HW-ID Fail"
             'ContFail = ContFail + 1
        
        Case WM_FT_TESTUNITREADY_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "HV: TestUnitReady Fail"
             'ContFail = ContFail + 1
        
        Case WM_FT_RW_SPEED_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "HV: SPEED Error "
             'ContFail = ContFail + 1

        Case WM_FT_RW_RW_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "HV: RW FAIL "
             'ContFail = ContFail + 1

        Case WM_FT_CHECK_CERBGPO_FAIL

            TestResult = "Bin3"
            MPTester.TestResultLab = "HV: GPO/RB FAIL "
            'ContFail = ContFail + 1

        Case WM_FT_RW_ROM_FAIL
            TestResult = "Bin4"
            MPTester.TestResultLab = "HV: ROM FAIL "
            'ContFail = ContFail + 1

        Case WM_FT_RW_RAM_FAIL, WM_FT_PHYREAD_FAIL, WM_FT_ECC_FAIL, WM_FT_NOFREEBLOCK_FAIL, WM_FT_LODECODE_FAIL, WM_FT_RELOADCODE_FAIL
            TestResult = "Bin5"
            MPTester.TestResultLab = "HV: RAM FAIL "
            'ContFail = ContFail + 1

        Case WM_FT_RW_RW_PASS

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
                'ContFail = 0 '
            Else

                TestResult = "Bin3"

            End If
                       
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = "HV: Undefine Fail"

             ContFail = ContFail + 1
        
               
        End Select
        
        HV_Result = TestResult
        TestResult = ""
        
        cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'Power OFF UNLoad Device
        Call MsecDelay(0.2)
        Call PowerSet2(1, "3.0", "0.15", 1, "3.0", "0.15", 1)
        EQC_LV = True
        cardresult = DO_WritePort(card, Channel_P1A, &HFA)  'Power ON
        Call MsecDelay(1.3)
        'WaitDevOn ("vid_058f")
        'Call MsecDelay(0.2)
        GoTo RW_Test_Label
        
    ElseIf (EQC_HV = True) And (EQC_LV = True) Then
        
        Select Case AlcorMPMessage
  
        Case WM_FT_RW_UNKNOW_FAIL
            TestResult = "Bin2"
            MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: UnKnow Fail"
            'ContFail = ContFail + 1
        
        Case WM_FT_CHECK_HW_CODE_FAIL
             TestResult = "Bin5"
             MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: HW-ID Fail"
             'ContFail = ContFail + 1
        
        Case WM_FT_TESTUNITREADY_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: TestUnitReady Fail"
             'ContFail = ContFail + 1
        
        Case WM_FT_RW_SPEED_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: SPEED Error "
             'ContFail = ContFail + 1

        Case WM_FT_RW_RW_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: RW FAIL "
             'ContFail = ContFail + 1

        Case WM_FT_CHECK_CERBGPO_FAIL

            TestResult = "Bin3"
            MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: GPO/RB FAIL "
            'ContFail = ContFail + 1

        Case WM_FT_RW_ROM_FAIL
            TestResult = "Bin4"
            MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: ROM FAIL "
            'ContFail = ContFail + 1

        Case WM_FT_RW_RAM_FAIL, WM_FT_PHYREAD_FAIL, WM_FT_ECC_FAIL, WM_FT_NOFREEBLOCK_FAIL, WM_FT_LODECODE_FAIL, WM_FT_RELOADCODE_FAIL
            TestResult = "Bin5"
            MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: RAM FAIL "
            'ContFail = ContFail + 1

        Case WM_FT_RW_RW_PASS

            For LedCount = 1 To 20
                Call MsecDelay(0.1)
                cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
                
                If LightOn = &HEF Or LightOn = &HCF Or LightOn = 223 Then
                    Exit For
                End If
            Next LedCount

            MPTester.Print "light="; LightOn

            If (LightOn = &HEF Or LightOn = &HCF Or LightOn = 223) Then
                MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: PASS"
                TestResult = "PASS"
                'ContFail = 0 '
            Else

                TestResult = "Bin3"

            End If
                       
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: Undefine Fail"

             'ContFail = ContFail + 1
        
               
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
        
        cardresult = DO_WritePort(card, Channel_P1A, &HFA)
    
    End If
        
         
                            
                            
End Sub

Public Sub AU6996_MP_Golden()

    If ChipName = "AU6996A51ILF2B" Then
        ChipName = "AU6996A51ILF2A"
    End If
    
    Call AU6996Flash_DualTestSub
    
End Sub


Public Sub AU6992A52HLS10SortingSub()
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
 
If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
    FileCopy App.Path & "\AlcorMP_698x2\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
    Call MsecDelay(5)
End If


NewChipFlag = 0
If OldChipName <> ChipName Then
            FileCopy App.Path & "\AlcorMP_698x2\ROM\" & ChipName & "\ROM.Hex", App.Path & "\AlcorMP_698x2\ROM.Hex"
            FileCopy App.Path & "\AlcorMP_698x2\RAM\" & ChipName & "\RAM.Bin", App.Path & "\AlcorMP_698x2\RAM.Bin"
            FileCopy App.Path & "\AlcorMP_698x2\INI\" & ChipName & "\AlcorMP.ini", App.Path & "\AlcorMP_698x2\AlcorMP.ini"
            NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName
 

'==============================================================
' when begin RW Test, must clear MP program
'===============================================================
'(1)
winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
  Loop While winHwnd <> 0
End If
winHwnd = FindWindow(vbNullString, AU6992MPCaption)
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, AU6992MPCaption)
  Loop While winHwnd <> 0
End If
'(2)
winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
  Loop While winHwnd <> 0
End If
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail
 
 
 '====================================
 '  Fix Card
 '====================================
' GoTo T1
 If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag = True) Then
 
 
   If MPTester.NoMP.Value = 1 Then
        
        If (NewChipFlag = 0) And (MPTester.Check1.Value = 0) Then  ' force condition
           GoTo RW_Test_Label
        End If
    End If
       
    
    If MPTester.ResetMPFailCounter.Value = 1 Then
       ContFail = 0
    End If
    
 '==============================================================
' when begin MP, must close RW program
'===============================================================
   MPFlag = 1
 
    winHwnd = FindWindow(vbNullString, "UFD Test")
    If winHwnd <> 0 Then
      Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "UFD Test")
      Loop While winHwnd <> 0
    End If
 
       '  power on
       cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        Call PowerSet2(2, "5.0", "0.5", 1, "3.3", "0.5", 1) ' close power to disable chip
       Call MsecDelay(0.5)  ' power for load MPDriver
       MPTester.Print "wait for MP Ready"
       Call LoadMP_Click_AU6992
 
        OldTimer = Timer
        AlcorMPMessage = 0
        Debug.Print "begin"
        Do
           ' DoEvents
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
            PassTime = Timer - OldTimer
            'Debug.Print AlcorMPMessage
        Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 30 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3
        '====================================================
        If PassTime > 30 Then
          '(1)
           MPTester.Print "MP Ready Fail"
            winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
              Loop While winHwnd <> 0
            End If
           '(2)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
            If winHwnd <> 0 Then
              Do
              rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
              Call MsecDelay(0.5)
              winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
              Loop While winHwnd <> 0
            End If
            
        
             MPTester.TestResultLab = "Bin3:MP Ready Fail"
             TestResult = "Bin3"
             MPTester.Print "MP Ready Fail"
     
              
            Exit Sub
        End If
        
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_MP_START Then
         
             
             cardresult = DO_WritePort(card, Channel_P1A, &HFB)  ' sel chip
              Call PowerSet2(1, "5.0", "0.5", 1, "3.3", "0.5", 1) ' close power to disable chip
             
 
          
             
            Dim TimerCounter As Integer
            Dim TmpString As String
            
             
            Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName("vid")
            Loop While TmpString = "" And TimerCounter < 150
             
             
             If TmpString = "" Then   ' can not find device after 15 s
             
               TestResult = "Bin2"
               MPTester.TestResultLab = "Bin2:MP UNKNOW Fail when enter MP"
               Exit Sub
             End If
             
             Call MsecDelay(2.5)
               
             MPTester.Print " MP Begin....."
             
             Call StartMP_Click_AU6992
   
             
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
                
               ' time out fail
                If PassTime > 65 Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    '(1)
                    winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                      Loop While winHwnd <> 0
                    End If
                    
                    '(2)
                      winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    Exit Sub
                End If
                
                ' MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    
                    winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                      Loop While winHwnd <> 0
                    End If
                    
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    
                    Exit Sub
                End If
                
                
                 'unknow fail
                 If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                     MPContFail = MPContFail + 1
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                    MPTester.Print "MP UNKNOW Fail"
                    
                   winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                    If winHwnd <> 0 Then
                     Do
                     rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                     Call MsecDelay(0.5)
                     winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                     Loop While winHwnd <> 0
                     End If
                     
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                     If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                     End If
                     
                     
                     
                    Exit Sub
                End If
                 
                
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_PASS Then
                     MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPTester.Print "MP PASS"
                End If
        End If
   
End If
'=========================================
 '    Close MP program
 '=========================================
winHwnd = FindWindow(vbNullString, AU6992MPCaption)
If winHwnd <> 0 Then
  Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, AU6992MPCaption)
  Loop While winHwnd <> 0
End If
    
 winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
    Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
    Loop While winHwnd <> 0
End If

  Dim pid As Long          ' unload driver
  Dim hProcess As Long
  Dim ExitEvent As Long
 
  pid = Shell(App.Path & "\AlcorMP_698x2\loaddrv.exe uninstall_058F6387")
  hProcess = OpenProcess(SYNCHRONIZE + PROCESS_QUERY_INFORMATION + PROCESS_TERMINATE, 0, pid)
  ExitEvent = WaitForSingleObject(hProcess, INFINITE)
  Call CloseHandle(hProcess)
  KillProcess ("LoadDrv.exe")
 
 
                        
 '=========================================
 '    POWER on
 '=========================================
'T1:
RW_Test_Label:


 
     If MPFlag = 1 Then
              Call PowerSet2(2, "5.0", "0.5", 1, "3.3", "0.5", 1) ' close power to disable chip
              cardresult = DO_WritePort(card, Channel_P1A, &HFF)
            
             Call MsecDelay(0.5)  ' power of to unload MPDriver
    
               cardresult = DO_WritePort(card, Channel_P1A, &HFB)  'sel socket
              Call PowerSet2(1, "5.0", "0.15", 1, "3.3", "0.15", 1) ' close power to disable chip
         
            
             Call MsecDelay(1.2)
            MPFlag = 0
     Else
              cardresult = DO_WritePort(card, Channel_P1A, &HFB)
             Call PowerSet2(1, "5.0", "0.15", 1, "3.3", "0.15", 1) ' close power to disable chi
             
             Call MsecDelay(1.2)
    End If
             Call LoadRWTest_Click_AU6992

        
        
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
     '   GoTo T2
       If PassTime > 5 Then
           TestResult = "Bin3"
           MPTester.TestResultLab = "Bin3:RW Ready Fail"
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
       
            Exit Sub
       End If
         
T2:
Dim Vol As Single
Dim AlcorMPMessageExp As String
        MPTester.Cls
       For Vol = 3.6 To 2.99 Step -0.05
                    
        Call GPIBWrite("VSET2 " & CStr(Vol))
        
                OldTimer = Timer
                AlcorMPMessage = 0
              '  MPTester.Print "RW Tester begin test........"
                Call StartRWTest_Click_AU6988
                
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
                        Or AlcorMPMessage = WM_FT_CHECK_CERBGPO_FAIL _
                      Or PassTime > 20 _
                      Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
            
                 ' MPTester.Print "RW work Time="; PassTime
                  MPTester.MPText.Text = Hex(AlcorMPMessage)
                
        
                 '===========================================================
                 '  RW Time Out Fail
                 '===========================================================
                 
                 If PassTime > 20 Then
                     TestResult = "Bin3"
                     MPTester.TestResultLab = "Bin3:RW Time Out Fail"
                   
                     winHwnd = FindWindow(vbNullString, "UFD Test")
                     If winHwnd <> 0 Then
                       Do
                         rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                         Call MsecDelay(0.5)
                         winHwnd = FindWindow(vbNullString, "UFD Test")
                       Loop While winHwnd <> 0
                     End If
                     
                       cardresult = DO_WritePort(card, Channel_P1A, &HFB)  ' power off
                     
                       Call PowerSet2(1, "5.0", "0.15", 1, "3.3", "0.15", 1) ' close power to disable chi
                 
                     
                
                     Exit Sub
                 End If
        
        
                   If AlcorMPMessage = WM_FT_RW_UNKNOW_FAIL Then
                   
                       AlcorMPMessageExp = "UNKNOW_FAIL"
                   ElseIf AlcorMPMessage = WM_FT_RW_SPEED_FAIL Then
                   
                       AlcorMPMessageExp = "SPEED_FAIL"
                       
                     ElseIf AlcorMPMessage = WM_FT_RW_RW_FAIL Then
                   
                       AlcorMPMessageExp = "RW_FAIL"
        
                    ElseIf AlcorMPMessage = WM_FT_CHECK_CERBGPO_FAIL Then
                   
                       AlcorMPMessageExp = "CERBGPO_FAIL"
                       
                    ElseIf AlcorMPMessage = WM_FT_RW_ROM_FAIL Then
                   
                       AlcorMPMessageExp = "ROM_FAIL"
                       
                      ElseIf AlcorMPMessage = WM_FT_RW_RAM_FAIL Then
                   
                       AlcorMPMessageExp = "RAM_FAIL"
                       
                     ElseIf AlcorMPMessage = WM_FT_RW_RW_PASS Then
                   
                       AlcorMPMessageExp = "PASS"
                     End If
                       
                       
                                           
        
                    MPTester.Print "Vol="; Vol; AlcorMPMessageExp; " Time="; PassTime
                    If AlcorMPMessage <> WM_FT_RW_RW_PASS Then
                       
                        Exit For
                    End If
            Next Vol
        
     
               
        Select Case AlcorMPMessage
        
        Case WM_FT_RW_UNKNOW_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:UnKnow Fail"
          
             ContFail = ContFail + 1
        
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
             
        Case WM_FT_RW_ROM_FAIL
              TestResult = "Bin4"
              MPTester.TestResultLab = "Bin4:ROM FAIL "
              ContFail = ContFail + 1
              
        Case WM_FT_RW_RAM_FAIL
              TestResult = "Bin5"
              MPTester.TestResultLab = "Bin5:RAM FAIL "
               ContFail = ContFail + 1
        Case WM_FT_RW_RW_PASS
        
               
              
                    MPTester.TestResultLab = "PASS "
                    TestResult = "PASS"
                    ContFail = 0
            
               
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:Undefine Fail"
          
             ContFail = ContFail + 1
        
               
        End Select
                               
       cardresult = DO_WritePort(card, Channel_P1A, &HFB)
           Call PowerSet2(1, "5.0", "0.15", 1, "3.3", "0.15", 1) ' close power to disable chi
                 
        
         
                            
End Sub

Public Sub AU6996A51BLF21TestSub()
'add unload driver function
' use AU6996 new MP and AU6992 RW test

    Lon = False
    Loff = False
    LightSituation = 255

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
'===============================================================3
 
If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
    FileCopy App.Path & "\AlcorMP_698x2\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
    Call MsecDelay(5)
End If


NewChipFlag = 0
If OldChipName <> ChipName Then
            FileCopy App.Path & "\AlcorMP_6996Flash\New_INI\" & ChipName & "\ROM\ROM.Hex", App.Path & "\AlcorMP_6996Flash\ROM.Hex"
            FileCopy App.Path & "\AlcorMP_6996Flash\New_INI\" & ChipName & "\RAM\RAM.Bin", App.Path & "\AlcorMP_6996Flash\RAM.Bin"
            FileCopy App.Path & "\AlcorMP_6996Flash\New_INI\" & ChipName & "\AlcorMP.ini", App.Path & "\AlcorMP_6996Flash\AlcorMP.ini"
            FileCopy App.Path & "\AlcorMP_6996Flash\New_INI\" & ChipName & "\PE.bin", App.Path & "\AlcorMP_6996Flash\PE.bin"
            NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName

SetSiteStatus (SiteReady)
 

'==============================================================
' when begin RW Test, must clear MP program
'===============================================================
'(1)
winHwnd = FindWindow(vbNullString, AU6996MPCaption1)
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, AU6996MPCaption1)
  Loop While winHwnd <> 0
End If


winHwnd = FindWindow(vbNullString, AU6996MPCaption)
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, AU6996MPCaption)
  Loop While winHwnd <> 0
End If
'(2)
winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
  Loop While winHwnd <> 0
End If
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail
 
 
 '====================================
 '  Fix Card
 '====================================
' GoTo T1
 If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag = True) Then
 
 
   If MPTester.NoMP.Value = 1 Then
        
        If (NewChipFlag = 0) And (MPTester.Check1.Value = 0) Then  ' force condition
           GoTo RW_Test_Label
        End If
    End If
       
    
    If MPTester.ResetMPFailCounter.Value = 1 Then
       ContFail = 0
    End If
    
 '==============================================================
' when begin MP, must close RW program
'===============================================================
   MPFlag = 1
 
    winHwnd = FindWindow(vbNullString, "UFD Test")
    If winHwnd <> 0 Then
      Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "UFD Test")
      Loop While winHwnd <> 0
    End If
 
       '  power on
       cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        Call PowerSet(3)   ' close power to disable chip
       Call MsecDelay(0.5)  ' power for load MPDriver
       MPTester.Print "wait for MP Ready"
       Call LoadNewMP_Click_AU6996Flash
 
        OldTimer = Timer
        AlcorMPMessage = 0
        Debug.Print "begin"
        Do
           ' DoEvents
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
            PassTime = Timer - OldTimer
            'Debug.Print AlcorMPMessage
        Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 30 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3
        '====================================================
        If PassTime > 30 Then
          '(1)
          SetSiteStatus (SiteUnknow)
           MPTester.Print "MP Ready Fail"
            winHwnd = FindWindow(vbNullString, AU6996MPCaption1)
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, AU6996MPCaption1)
              Loop While winHwnd <> 0
            End If
           '(2)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
            If winHwnd <> 0 Then
              Do
              rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
              Call MsecDelay(0.5)
              winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
              Loop While winHwnd <> 0
            End If
            
        
             MPTester.TestResultLab = "Bin3:MP Ready Fail"
             TestResult = "Bin3"
             MPTester.Print "MP Ready Fail"
     
              
            Exit Sub
        End If
        
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_MP_START Then
            SetSiteStatus (RunMP)
            WaitAnotherSiteDone (RunMP)
             cardresult = DO_WritePort(card, Channel_P1A, &HFB)  ' sel chip
              Call PowerSet(500)   ' close power to disable chip
             
 
          
             
            Dim TimerCounter As Integer
            Dim TmpString As String
            
             
            Do
            DoEvents
                Call MsecDelay(0.1)
                TimerCounter = TimerCounter + 1
                TmpString = GetDeviceName("vid")
            Loop While TmpString = "" And TimerCounter < 150
             
             
             If TmpString = "" Then   ' can not find device after 15 s
             
               TestResult = "Bin2"
               MPTester.TestResultLab = "Bin2:MP UNKNOW Fail when enter MP"
               Exit Sub
             End If
             
             Call MsecDelay(3)
               
             MPTester.Print " MP Begin....."
             
             Call StartMP_Click_AU6996
             
             Call MsecDelay(1)
   
             
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
                
               ' time out fail
                If PassTime > 65 Then
                    SetSiteStatus (SiteUnknow)
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    '(1)
                    winHwnd = FindWindow(vbNullString, AU6996MPCaption)
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, AU6996MPCaption)
                      Loop While winHwnd <> 0
                    End If
                    
                    '(2)
                      winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    Exit Sub
                End If
                
                ' MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL Then
                    SetSiteStatus (SiteUnknow)
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    
                    winHwnd = FindWindow(vbNullString, AU6996MPCaption)
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, AU6996MPCaption)
                      Loop While winHwnd <> 0
                    End If
                    
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    
                    Exit Sub
                End If
                
                
                 'unknow fail
                 If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                    SetSiteStatus (SiteUnknow)
                     MPContFail = MPContFail + 1
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                    MPTester.Print "MP UNKNOW Fail"
                    
                   winHwnd = FindWindow(vbNullString, AU6996MPCaption)
                    If winHwnd <> 0 Then
                     Do
                     rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                     Call MsecDelay(0.5)
                     winHwnd = FindWindow(vbNullString, AU6996MPCaption)
                     Loop While winHwnd <> 0
                     End If
                     
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                     If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                     End If
                     
                     
                     
                    Exit Sub
                End If
                 
                
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_PASS Then
                    SetSiteStatus (MPDone)
                     MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPTester.Print "MP PASS"
                End If
        End If
   
End If
'=========================================
 '    Close MP program
 '=========================================
winHwnd = FindWindow(vbNullString, AU6996MPCaption)
If winHwnd <> 0 Then
  Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, AU6996MPCaption)
  Loop While winHwnd <> 0
End If
    
 winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
    Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
    Loop While winHwnd <> 0
End If

  Dim pid As Long          ' unload driver
  Dim hProcess As Long
  Dim ExitEvent As Long
 
  pid = Shell(App.Path & "\AlcorMP_6996Flash\loaddrv.exe uninstall_058F6387")
  hProcess = OpenProcess(SYNCHRONIZE + PROCESS_QUERY_INFORMATION + PROCESS_TERMINATE, 0, pid)
  ExitEvent = WaitForSingleObject(hProcess, INFINITE)
  Call CloseHandle(hProcess)
  KillProcess ("LoadDrv.exe")
 
 
                        
 '=========================================
 '    POWER on
 '=========================================
'T1:
RW_Test_Label:
 
 If MPFlag = 1 Then
        WaitAnotherSiteDone (MPDone)
         Call PowerSet(3)
          cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        
         Call MsecDelay(0.5)  ' power of to unload MPDriver

           cardresult = DO_WritePort(card, Channel_P1A, &HFB)  'sel socket
           Call PowerSet2(1, "5.0", "0.2", 1, "5.0", "0.2", 1)
     
        
         WaitDevOn ("058f")
        MPFlag = 0
 Else
        MsecDelay (0.1)
          cardresult = DO_WritePort(card, Channel_P1A, &HFB)
         Call PowerSet2(1, "5.0", "0.2", 1, "5.0", "0.2", 1)
         
         WaitDevOn ("058f")
End If
         Call LoadRWTest_Click_AU6996Flash
         
'                 ' 20131129 add for led check
'        For LedCount = 1 To 20
'            Call MsecDelay(0.1)
'            cardresult = DO_ReadPort(card, Channel_P1B, LightSituation)
'            If LightSituation = 255 Then
'                Loff = True
'            Else
'                Lon = True
'            End If
'
'            If (Loff = True) And (Lon = True) Then
'                Exit For
'            End If
'
'        Next LedCount
        
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
     '   GoTo T2
       If PassTime > 5 Then
           TestResult = "Bin3"
           MPTester.TestResultLab = "Bin3:RW Ready Fail"
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
       
            Exit Sub
       End If
         
T2:
        SetSiteStatus (RunHV)
        WaitAnotherSiteDone (RunHV)

        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        Call StartRWTest_Click_AU6988
        
        MsecDelay (0.1)
        
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
                Or AlcorMPMessage = WM_FT_CHECK_CERBGPO_FAIL _
                  Or AlcorMPMessage = WM_FT_PARAM_FAIL _
                  Or AlcorMPMessage = WM_FT_READER_FAIL _
                  Or AlcorMPMessage = WM_FT_BUSWIDTH_FAIL _
                  Or AlcorMPMessage = WM_FT_BUSCLK_FAIL _
              Or PassTime > 5 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
              MsecDelay (0.1)
    
        ' 20131129 add for led check
        For LedCount = 1 To 20
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
    
          MPTester.Print "RW work Time="; PassTime
          MPTester.MPText.Text = Hex(AlcorMPMessage)
        
        
        '===========================================================
        '  RW Time Out Fail
        '===========================================================
        
        If (PassTime > 5) Or ((FailCloseAP) And (AlcorMPMessage <> WM_FT_RW_RW_PASS)) Then
            Close_FT_AP ("UFD Test")
            SetSiteStatus (SiteUnknow)
            If (PassTime > 5) Then
                TestResult = "Bin3"
                MPTester.TestResultLab = "Bin3:RW Time Out Fail"
                cardresult = DO_WritePort(card, Channel_P1A, &HFB)  ' power off
                Exit Sub
            End If
        
        End If
     
               
        Select Case AlcorMPMessage
        
        Case WM_FT_RW_UNKNOW_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:UnKnow Fail"
          
             ContFail = ContFail + 1
        
        Case WM_FT_RW_SPEED_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:SPEED Error "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_RW_FAIL, WM_FT_PARAM_FAIL, WM_FT_READER_FAIL, WM_FT_BUSWIDTH_FAIL, WM_FT_BUSCLK_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:RW FAIL "
             ContFail = ContFail + 1
             
        Case WM_FT_CHECK_CERBGPO_FAIL
        
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:GPO/RB FAIL "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_ROM_FAIL
              TestResult = "Bin4"
              MPTester.TestResultLab = "Bin4:ROM FAIL "
              ContFail = ContFail + 1
              
        Case WM_FT_RW_RAM_FAIL
              TestResult = "Bin5"
              MPTester.TestResultLab = "Bin5:RAM FAIL "
               ContFail = ContFail + 1
        Case WM_FT_RW_RW_PASS
        
               
               For LedCount = 1 To 20
               Call MsecDelay(0.1)
               cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
                If LightOn = &HEF Or LightOn = &HCF Or LightOn = 223 Then
               
                 Exit For
               End If
               Next LedCount
                 
                ' 20131129 add check led off
                 MPTester.Print "light="; LightOn
                 If (LightOn = &HEF Or LightOn = &HCF Or LightOn = 223) And (Loff = True) Then
                    MPTester.TestResultLab = "PASS "
                    TestResult = "PASS"
                    ContFail = 0
                Else
                 
                  TestResult = "Bin3"
                  MPTester.TestResultLab = "Bin3:LED FAIL "
              
               End If
               
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:Undefine Fail"
          
             ContFail = ContFail + 1
        
               
        End Select
                               
       cardresult = DO_WritePort(card, Channel_P1A, &HFB)
         'Call PowerSet(1500)
        Call PowerSet2(1, "0.0", "0.2", 1, "0.0", "0.2", 1)
        WaitDevOFF ("058f")
                                    
End Sub

Public Sub AU6996A51BLF31TestSub()
'add unload driver function
' use AU6996 new MP and AU6992 RW test

    Lon = False
    Loff = False
    LightSituation = 255

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
'===============================================================3
 
If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
    FileCopy App.Path & "\AlcorMP_698x2\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
    Call MsecDelay(5)
End If


NewChipFlag = 0
If OldChipName <> ChipName Then
            FileCopy App.Path & "\AlcorMP_6996Flash\New_INI\" & ChipName & "\ROM\ROM.Hex", App.Path & "\AlcorMP_6996Flash\ROM.Hex"
            FileCopy App.Path & "\AlcorMP_6996Flash\New_INI\" & ChipName & "\RAM\RAM.Bin", App.Path & "\AlcorMP_6996Flash\RAM.Bin"
            FileCopy App.Path & "\AlcorMP_6996Flash\New_INI\" & ChipName & "\AlcorMP.ini", App.Path & "\AlcorMP_6996Flash\AlcorMP.ini"
            FileCopy App.Path & "\AlcorMP_6996Flash\New_INI\" & ChipName & "\PE.bin", App.Path & "\AlcorMP_6996Flash\PE.bin"
            NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName
 

'==============================================================
' when begin RW Test, must clear MP program
'===============================================================
'(1)
winHwnd = FindWindow(vbNullString, AU6996MPCaption1)
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, AU6996MPCaption1)
  Loop While winHwnd <> 0
End If


winHwnd = FindWindow(vbNullString, AU6996MPCaption)
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, AU6996MPCaption)
  Loop While winHwnd <> 0
End If
'(2)
winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
  Loop While winHwnd <> 0
End If
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail
 
 
 '====================================
 '  Fix Card
 '====================================
' GoTo T1
 If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag = True) Then
 
 
   If MPTester.NoMP.Value = 1 Then
        
        If (NewChipFlag = 0) And (MPTester.Check1.Value = 0) Then  ' force condition
           GoTo RW_Test_Label
        End If
    End If
       
    
    If MPTester.ResetMPFailCounter.Value = 1 Then
       ContFail = 0
    End If
    
 '==============================================================
' when begin MP, must close RW program
'===============================================================
   MPFlag = 1
 
    winHwnd = FindWindow(vbNullString, "UFD Test")
    If winHwnd <> 0 Then
      Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "UFD Test")
      Loop While winHwnd <> 0
    End If
 
       '  power on
       cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        Call PowerSet(3)   ' close power to disable chip
       Call MsecDelay(0.5)  ' power for load MPDriver
       MPTester.Print "wait for MP Ready"
       Call LoadNewMP_Click_AU6996Flash
 
        OldTimer = Timer
        AlcorMPMessage = 0
        Debug.Print "begin"
        Do
           ' DoEvents
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
            PassTime = Timer - OldTimer
            'Debug.Print AlcorMPMessage
        Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 30 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3
        '====================================================
        If PassTime > 30 Then
          '(1)
           MPTester.Print "MP Ready Fail"
            winHwnd = FindWindow(vbNullString, AU6996MPCaption1)
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, AU6996MPCaption1)
              Loop While winHwnd <> 0
            End If
           '(2)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
            If winHwnd <> 0 Then
              Do
              rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
              Call MsecDelay(0.5)
              winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
              Loop While winHwnd <> 0
            End If
            
        
             MPTester.TestResultLab = "Bin3:MP Ready Fail"
             TestResult = "Bin3"
             MPTester.Print "MP Ready Fail"
     
              
            Exit Sub
        End If
        
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_MP_START Then
         
             
             cardresult = DO_WritePort(card, Channel_P1A, &HFB)  ' sel chip
             'Call PowerSet(500)   ' close power to disable chip
              
             Call PowerSet2(1, "3.3", "0.2", 1, "3.3", "0.2", 1)
 
          
             
            Dim TimerCounter As Integer
            Dim TmpString As String
            
             
            Do
            DoEvents
                Call MsecDelay(0.1)
                TimerCounter = TimerCounter + 1
                TmpString = GetDeviceName("vid")
            Loop While TmpString = "" And TimerCounter < 150
             
             
             If TmpString = "" Then   ' can not find device after 15 s
             
               TestResult = "Bin2"
               MPTester.TestResultLab = "Bin2:MP UNKNOW Fail when enter MP"
               Exit Sub
             End If
             
             Call MsecDelay(2.5)
               
             MPTester.Print " MP Begin....."
             
             Call StartMP_Click_AU6996
   
             
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
                
               ' time out fail
                If PassTime > 65 Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    '(1)
                    winHwnd = FindWindow(vbNullString, AU6996MPCaption)
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, AU6996MPCaption)
                      Loop While winHwnd <> 0
                    End If
                    
                    '(2)
                      winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    Exit Sub
                End If
                
                ' MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL Then
                    MPFlag = 1
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    
                    winHwnd = FindWindow(vbNullString, AU6996MPCaption)
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, AU6996MPCaption)
                      Loop While winHwnd <> 0
                    End If
                    
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    
                    Exit Sub
                End If
                
                
                 'unknow fail
                 If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                     MPContFail = MPContFail + 1
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                    MPTester.Print "MP UNKNOW Fail"
                    
                   winHwnd = FindWindow(vbNullString, AU6996MPCaption)
                    If winHwnd <> 0 Then
                     Do
                     rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                     Call MsecDelay(0.5)
                     winHwnd = FindWindow(vbNullString, AU6996MPCaption)
                     Loop While winHwnd <> 0
                     End If
                     
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                     If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                     End If
                     
                     
                     
                    Exit Sub
                End If
                 
                
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_PASS Then
                     MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPTester.Print "MP PASS"
                End If
        End If
   
End If
'=========================================
 '    Close MP program
 '=========================================
winHwnd = FindWindow(vbNullString, AU6996MPCaption)
If winHwnd <> 0 Then
  Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, AU6996MPCaption)
  Loop While winHwnd <> 0
End If
    
 winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
    Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
    Loop While winHwnd <> 0
End If

  Dim pid As Long          ' unload driver
  Dim hProcess As Long
  Dim ExitEvent As Long
 
  pid = Shell(App.Path & "\AlcorMP_6996Flash\loaddrv.exe uninstall_058F6387")
  hProcess = OpenProcess(SYNCHRONIZE + PROCESS_QUERY_INFORMATION + PROCESS_TERMINATE, 0, pid)
  ExitEvent = WaitForSingleObject(hProcess, INFINITE)
  Call CloseHandle(hProcess)
  KillProcess ("LoadDrv.exe")
 
 EQC_HV = False
 EQC_LV = False
                        
 '=========================================
 '    POWER on
 '=========================================
'T1:
RW_Test_Label:
 
    If MPFlag = 1 Then
    
        Call PowerSet(3)
        cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        Call MsecDelay(0.5)  ' power of to unload MPDriver
        
        cardresult = DO_WritePort(card, Channel_P1A, &HFB)  'sel socket
        Call PowerSet2(1, "3.6", "0.2", 1, "3.6", "0.2", 1)
        Call MsecDelay(3.2)
        MPFlag = 0
        EQC_HV = True
    Else
        cardresult = DO_WritePort(card, Channel_P1A, &HFB)
         
        If (EQC_HV = False) And (EQC_LV = False) Then
           Call PowerSet2(1, "3.6", "0.2", 1, "3.6", "0.2", 1)
           EQC_HV = True
        Else
           Call PowerSet2(1, "3.0", "0.2", 1, "3.0", "0.2", 1)
           EQC_LV = True
        End If
        
        Call MsecDelay(1.7)
    End If
    
    Call LoadRWTest_Click_AU6996Flash
         
        
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
        Loop Until AlcorMPMessage = WM_FT_RW_READY _
                Or AlcorMPMessage = WM_CLOSE _
                Or AlcorMPMessage = WM_DESTROY _
                Or PassTime > 5

        MPTester.Print "RW Ready Time="; PassTime
        
     '   GoTo T2
        If PassTime > 5 Then
            TestResult = "Bin3"
            MPTester.TestResultLab = "Bin3:RW Ready Fail"
        
            winHwnd = FindWindow(vbNullString, "UFD Test")
            
            If winHwnd <> 0 Then
                Do
                    rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                    Call MsecDelay(0.5)
                    winHwnd = FindWindow(vbNullString, "UFD Test")
                Loop While winHwnd <> 0
            End If
            
            Exit Sub
        End If

T2:
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        Call StartRWTest_Click_AU6988
        
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
                Or AlcorMPMessage = WM_FT_CHECK_CERBGPO_FAIL _
                  Or AlcorMPMessage = WM_FT_PARAM_FAIL _
                  Or AlcorMPMessage = WM_FT_READER_FAIL _
                  Or AlcorMPMessage = WM_FT_BUSWIDTH_FAIL _
                  Or AlcorMPMessage = WM_FT_BUSCLK_FAIL _
              Or PassTime > 5 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
    
        ' 20131129 add for led check
        For LedCount = 1 To 20
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
    
          MPTester.Print "RW work Time="; PassTime
          MPTester.MPText.Text = Hex(AlcorMPMessage)
        
        
        '===========================================================
        '  RW Time Out Fail
        '===========================================================
        
        If (PassTime > 5) Or ((FailCloseAP) And (AlcorMPMessage <> WM_FT_RW_RW_PASS)) Then
            Close_FT_AP ("UFD Test")
            
            If (PassTime > 5) Then
                TestResult = "Bin3"
                MPTester.TestResultLab = "Bin3:RW Time Out Fail"
                cardresult = DO_WritePort(card, Channel_P1A, &HFB)  ' power off
                Exit Sub
            End If
        
        End If
     
               
'        Select Case AlcorMPMessage
'
'        Case WM_FT_RW_UNKNOW_FAIL
'             TestResult = "Bin2"
'             MPTester.TestResultLab = "Bin2:UnKnow Fail"
'
'             ContFail = ContFail + 1
'
'        Case WM_FT_RW_SPEED_FAIL
'             TestResult = "Bin3"
'             MPTester.TestResultLab = "Bin3:SPEED Error "
'             ContFail = ContFail + 1
'
'        Case WM_FT_RW_RW_FAIL, WM_FT_PARAM_FAIL, WM_FT_READER_FAIL, WM_FT_BUSWIDTH_FAIL, WM_FT_BUSCLK_FAIL
'             TestResult = "Bin3"
'             MPTester.TestResultLab = "Bin3:RW FAIL "
'             ContFail = ContFail + 1
'
'        Case WM_FT_CHECK_CERBGPO_FAIL
'
'             TestResult = "Bin3"
'             MPTester.TestResultLab = "Bin3:GPO/RB FAIL "
'             ContFail = ContFail + 1
'
'        Case WM_FT_RW_ROM_FAIL
'              TestResult = "Bin4"
'              MPTester.TestResultLab = "Bin4:ROM FAIL "
'              ContFail = ContFail + 1
'
'        Case WM_FT_RW_RAM_FAIL
'              TestResult = "Bin5"
'              MPTester.TestResultLab = "Bin5:RAM FAIL "
'               ContFail = ContFail + 1
'        Case WM_FT_RW_RW_PASS
'
'
'               For LedCount = 1 To 20
'               Call MsecDelay(0.1)
'               cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
'                If LightOn = &HEF Or LightOn = &HCF Or LightOn = 223 Then
'
'                 Exit For
'               End If
'               Next LedCount
'
'                ' 20131129 add check led off
'                 MPTester.Print "light="; LightOn
'                 If (LightOn = &HEF Or LightOn = &HCF Or LightOn = 223) And (Loff = True) Then
'                    MPTester.TestResultLab = "PASS "
'                    TestResult = "PASS"
'                    ContFail = 0
'                Else
'
'                  TestResult = "Bin3"
'                  MPTester.TestResultLab = "Bin3:LED FAIL "
'
'               End If
'
'        Case Else
'             TestResult = "Bin2"
'             MPTester.TestResultLab = "Bin2:Undefine Fail"
'
'             ContFail = ContFail + 1
'
'
'        End Select

        If (EQC_HV = True) And (EQC_LV = False) Then
                   
            Select Case AlcorMPMessage
      
            Case WM_FT_RW_UNKNOW_FAIL
                TestResult = "Bin2"
                MPTester.TestResultLab = "HV: UnKnow Fail"
                'ContFail = ContFail + 1
            
            Case WM_FT_CHECK_HW_CODE_FAIL
                 TestResult = "Bin5"
                 MPTester.TestResultLab = "HV: HW-ID Fail"
                 'ContFail = ContFail + 1
            
            Case WM_FT_TESTUNITREADY_FAIL
                 TestResult = "Bin2"
                 MPTester.TestResultLab = "HV: TestUnitReady Fail"
                 'ContFail = ContFail + 1
            
            Case WM_FT_RW_SPEED_FAIL
                 TestResult = "Bin3"
                 MPTester.TestResultLab = "HV: SPEED Error "
                 'ContFail = ContFail + 1
    
            Case WM_FT_RW_RW_FAIL
                 TestResult = "Bin3"
                 MPTester.TestResultLab = "HV: RW FAIL "
                 'ContFail = ContFail + 1
    
            Case WM_FT_CHECK_CERBGPO_FAIL
    
                TestResult = "Bin3"
                MPTester.TestResultLab = "HV: GPO/RB FAIL "
                'ContFail = ContFail + 1
    
            Case WM_FT_RW_ROM_FAIL
                TestResult = "Bin4"
                MPTester.TestResultLab = "HV: ROM FAIL "
                'ContFail = ContFail + 1
    
            Case WM_FT_RW_RAM_FAIL, WM_FT_PHYREAD_FAIL, WM_FT_ECC_FAIL, WM_FT_NOFREEBLOCK_FAIL, WM_FT_LODECODE_FAIL, WM_FT_RELOADCODE_FAIL
                TestResult = "Bin5"
                MPTester.TestResultLab = "HV: RAM FAIL "
                'ContFail = ContFail + 1
    
            Case WM_FT_RW_RW_PASS
    
                For LedCount = 1 To 20
                    Call MsecDelay(0.1)
                    cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
                    
                    If ((LightOn = &HEF Or LightOn = &HCF Or LightOn = 223) And (Loff = True)) Then
                        Exit For
                    End If

                Next LedCount
    
                MPTester.Print "light="; LightOn
                
                If ((LightOn = &HEF Or LightOn = &HCF Or LightOn = 223) And (Loff = True)) Then
                    MPTester.TestResultLab = "HV: PASS"
                    TestResult = "PASS"
                Else
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "HV: LED FAIL "
                End If
                           
            Case Else
                 TestResult = "Bin2"
                 MPTester.TestResultLab = "HV: Undefine Fail"
    
                 ContFail = ContFail + 1

            End Select
            
            HV_Result = TestResult
            TestResult = ""
            EQC_LV = True
            GoTo RW_Test_Label
            
        ElseIf (EQC_HV = True) And (EQC_LV = True) Then
        
            Select Case AlcorMPMessage
      
            Case WM_FT_RW_UNKNOW_FAIL
                TestResult = "Bin2"
                MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: UnKnow Fail"
                'ContFail = ContFail + 1
            
            Case WM_FT_CHECK_HW_CODE_FAIL
                 TestResult = "Bin5"
                 MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: HW-ID Fail"
                 'ContFail = ContFail + 1
            
            Case WM_FT_TESTUNITREADY_FAIL
                 TestResult = "Bin2"
                 MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: TestUnitReady Fail"
                 'ContFail = ContFail + 1
            
            Case WM_FT_RW_SPEED_FAIL
                 TestResult = "Bin3"
                 MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: SPEED Error "
                 'ContFail = ContFail + 1
    
            Case WM_FT_RW_RW_FAIL
                 TestResult = "Bin3"
                 MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: RW FAIL "
                 'ContFail = ContFail + 1
    
            Case WM_FT_CHECK_CERBGPO_FAIL
    
                TestResult = "Bin3"
                MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: GPO/RB FAIL "
                'ContFail = ContFail + 1
    
            Case WM_FT_RW_ROM_FAIL
                TestResult = "Bin4"
                MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: ROM FAIL "
                'ContFail = ContFail + 1
    
            Case WM_FT_RW_RAM_FAIL, WM_FT_PHYREAD_FAIL, WM_FT_ECC_FAIL, WM_FT_NOFREEBLOCK_FAIL, WM_FT_LODECODE_FAIL, WM_FT_RELOADCODE_FAIL
                TestResult = "Bin5"
                MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: RAM FAIL "
                'ContFail = ContFail + 1
    
            Case WM_FT_RW_RW_PASS
    
                For LedCount = 1 To 20
                    Call MsecDelay(0.1)
                    cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
                    
                    If ((LightOn = &HEF Or LightOn = &HCF Or LightOn = 223) And (Loff = True)) Then
                        Exit For
                    End If
                Next
    
                MPTester.Print "light="; LightOn
    
                If ((LightOn = &HEF Or LightOn = &HCF Or LightOn = 223) And (Loff = True)) Then
                    MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: PASS"
                    TestResult = "PASS"
                    'ContFail = 0 '
                Else
                    MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: LED FAIL "
                    TestResult = "Bin3"
    
                End If
                           
            Case Else
                 TestResult = "Bin2"
                 MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: Undefine Fail"
    
            End Select
        
        End If
        
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


        EQC_HV = False
        EQC_LV = False
                               
        cardresult = DO_WritePort(card, Channel_P1A, &HFB)
        Call PowerSet(3)
        'Call PowerSet2(1, "0.0", "0.2", 1, "0.0", "0.2", 1)
                                    
End Sub





Public Sub AU6996A51BLF20TestSub()
'add unload driver function
' use AU6996 new MP and AU6992 RW test

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
'===============================================================3
 
If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
    FileCopy App.Path & "\AlcorMP_698x2\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
    Call MsecDelay(5)
End If


NewChipFlag = 0
If OldChipName <> ChipName Then
            FileCopy App.Path & "\AlcorMP_6996Reader\PE_INI\" & ChipName & "\PE.ini", App.Path & "\AlcorMP_6996Reader\PE.ini"
            FileCopy App.Path & "\AlcorMP_6996Reader\ROM\" & ChipName & "\ROM.Hex", App.Path & "\AlcorMP_6996Reader\ROM.Hex"
            FileCopy App.Path & "\AlcorMP_6996Reader\RAM\" & ChipName & "\RAM.Bin", App.Path & "\AlcorMP_6996Reader\RAM.Bin"
            FileCopy App.Path & "\AlcorMP_6996Reader\INI\" & ChipName & "\AlcorMP.ini", App.Path & "\AlcorMP_6996Reader\AlcorMP.ini"
            NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName
 

'==============================================================
' when begin RW Test, must clear MP program
'===============================================================
'(1)
winHwnd = FindWindow(vbNullString, AU6996MPCaption1)
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, AU6996MPCaption1)
  Loop While winHwnd <> 0
End If


winHwnd = FindWindow(vbNullString, AU6996MPCaption)
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, AU6996MPCaption)
  Loop While winHwnd <> 0
End If
'(2)
winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
  Loop While winHwnd <> 0
End If
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail
 
 
 '====================================
 '  Fix Card
 '====================================
' GoTo T1
 If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag = True) Then
 
 
   If MPTester.NoMP.Value = 1 Then
        
        If (NewChipFlag = 0) And (MPTester.Check1.Value = 0) Then  ' force condition
           GoTo RW_Test_Label
        End If
    End If
       
    
    If MPTester.ResetMPFailCounter.Value = 1 Then
       ContFail = 0
    End If
    
 '==============================================================
' when begin MP, must close RW program
'===============================================================
   MPFlag = 1
 
    winHwnd = FindWindow(vbNullString, "UFD Test")
    If winHwnd <> 0 Then
      Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "UFD Test")
      Loop While winHwnd <> 0
    End If
 
       '  power on
       cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        Call PowerSet(3)   ' close power to disable chip
       Call MsecDelay(0.5)  ' power for load MPDriver
       MPTester.Print "wait for MP Ready"
       Call LoadMP_Click_AU6996Reader
 
        OldTimer = Timer
        AlcorMPMessage = 0
        Debug.Print "begin"
        Do
           ' DoEvents
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
            PassTime = Timer - OldTimer
            'Debug.Print AlcorMPMessage
        Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 30 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3
        '====================================================
        If PassTime > 30 Then
          '(1)
           MPTester.Print "MP Ready Fail"
            winHwnd = FindWindow(vbNullString, AU6996MPCaption1)
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, AU6996MPCaption1)
              Loop While winHwnd <> 0
            End If
           '(2)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
            If winHwnd <> 0 Then
              Do
              rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
              Call MsecDelay(0.5)
              winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
              Loop While winHwnd <> 0
            End If
            
        
             MPTester.TestResultLab = "Bin3:MP Ready Fail"
             TestResult = "Bin3"
             MPTester.Print "MP Ready Fail"
     
              
            Exit Sub
        End If
        
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_MP_START Then
         
             
             cardresult = DO_WritePort(card, Channel_P1A, &HFB)  ' sel chip
              Call PowerSet(500)   ' close power to disable chip
             
 
          
             
            Dim TimerCounter As Integer
            Dim TmpString As String
            
             
            Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName("vid")
            Loop While TmpString = "" And TimerCounter < 150
             
             
             If TmpString = "" Then   ' can not find device after 15 s
             
               TestResult = "Bin2"
               MPTester.TestResultLab = "Bin2:MP UNKNOW Fail when enter MP"
               Exit Sub
             End If
             
             Call MsecDelay(2.5)
               
             MPTester.Print " MP Begin....."
             
             Call StartMP_Click_AU6996
   
             
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
                
               ' time out fail
                If PassTime > 65 Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    '(1)
                    winHwnd = FindWindow(vbNullString, AU6996MPCaption)
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, AU6996MPCaption)
                      Loop While winHwnd <> 0
                    End If
                    
                    '(2)
                      winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    Exit Sub
                End If
                
                ' MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    
                    winHwnd = FindWindow(vbNullString, AU6996MPCaption)
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, AU6996MPCaption)
                      Loop While winHwnd <> 0
                    End If
                    
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    
                    Exit Sub
                End If
                
                
                 'unknow fail
                 If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                     MPContFail = MPContFail + 1
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                    MPTester.Print "MP UNKNOW Fail"
                    
                   winHwnd = FindWindow(vbNullString, AU6996MPCaption)
                    If winHwnd <> 0 Then
                     Do
                     rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                     Call MsecDelay(0.5)
                     winHwnd = FindWindow(vbNullString, AU6996MPCaption)
                     Loop While winHwnd <> 0
                     End If
                     
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                     If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                     End If
                     
                     
                     
                    Exit Sub
                End If
                 
                
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_PASS Then
                     MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPTester.Print "MP PASS"
                End If
        End If
   
End If
'=========================================
 '    Close MP program
 '=========================================
winHwnd = FindWindow(vbNullString, AU6996MPCaption)
If winHwnd <> 0 Then
  Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, AU6996MPCaption)
  Loop While winHwnd <> 0
End If
    
 winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
    Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
    Loop While winHwnd <> 0
End If

  Dim pid As Long          ' unload driver
  Dim hProcess As Long
  Dim ExitEvent As Long
 
  pid = Shell(App.Path & "\AlcorMP_6996Reader\loaddrv.exe uninstall_058F6387")
  hProcess = OpenProcess(SYNCHRONIZE + PROCESS_QUERY_INFORMATION + PROCESS_TERMINATE, 0, pid)
  ExitEvent = WaitForSingleObject(hProcess, INFINITE)
  Call CloseHandle(hProcess)
  KillProcess ("LoadDrv.exe")
 
 
                        
 '=========================================
 '    POWER on
 '=========================================
'T1:
RW_Test_Label:
 
 If MPFlag = 1 Then
         Call PowerSet(3)
          cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        
         Call MsecDelay(0.5)  ' power of to unload MPDriver

           cardresult = DO_WritePort(card, Channel_P1A, &HFB)  'sel socket
           Call PowerSet2(1, "5.0", "0.2", 1, "5.0", "0.2", 1)
     
        
         Call MsecDelay(3.2)
        MPFlag = 0
 Else
          cardresult = DO_WritePort(card, Channel_P1A, &HFB)
         Call PowerSet2(1, "5.0", "0.2", 1, "5.0", "0.2", 1)
         
         Call MsecDelay(1.7)
End If
         Call LoadRWTest_Click_AU6996Reader

        
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
     '   GoTo T2
       If PassTime > 5 Then
           TestResult = "Bin3"
           MPTester.TestResultLab = "Bin3:RW Ready Fail"
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
       
            Exit Sub
       End If
         
T2:
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        Call StartRWTest_Click_AU6988
        
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
                Or AlcorMPMessage = WM_FT_CHECK_CERBGPO_FAIL _
                  Or AlcorMPMessage = WM_FT_PARAM_FAIL _
                  Or AlcorMPMessage = WM_FT_READER_FAIL _
                  Or AlcorMPMessage = WM_FT_BUSWIDTH_FAIL _
                  Or AlcorMPMessage = WM_FT_BUSCLK_FAIL _
              Or PassTime > 5 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
    
          MPTester.Print "RW work Time="; PassTime
          MPTester.MPText.Text = Hex(AlcorMPMessage)
        
        
        '===========================================================
        '  RW Time Out Fail
        '===========================================================
        
        If (PassTime > 5) Or ((FailCloseAP) And (AlcorMPMessage <> WM_FT_RW_RW_PASS)) Then
            Close_FT_AP ("UFD Test")
            
            If (PassTime > 5) Then
                TestResult = "Bin3"
                MPTester.TestResultLab = "Bin3:RW Time Out Fail"
                cardresult = DO_WritePort(card, Channel_P1A, &HFB)  ' power off
                Exit Sub
            End If
        
        End If
     
               
        Select Case AlcorMPMessage
        
        Case WM_FT_RW_UNKNOW_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:UnKnow Fail"
          
             ContFail = ContFail + 1
        
        Case WM_FT_RW_SPEED_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:SPEED Error "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_RW_FAIL, WM_FT_PARAM_FAIL, WM_FT_READER_FAIL, WM_FT_BUSWIDTH_FAIL, WM_FT_BUSCLK_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:RW FAIL "
             ContFail = ContFail + 1
             
        Case WM_FT_CHECK_CERBGPO_FAIL
        
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:GPO/RB FAIL "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_ROM_FAIL
              TestResult = "Bin4"
              MPTester.TestResultLab = "Bin4:ROM FAIL "
              ContFail = ContFail + 1
              
        Case WM_FT_RW_RAM_FAIL
              TestResult = "Bin5"
              MPTester.TestResultLab = "Bin5:RAM FAIL "
               ContFail = ContFail + 1
        Case WM_FT_RW_RW_PASS
        
               
               For LedCount = 1 To 20
               Call MsecDelay(0.1)
               cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
                If LightOn = &HEF Or LightOn = &HCF Or LightOn = 223 Then
               
                 Exit For
               End If
               Next LedCount
                 
                ' 20131129 add check led off
                 MPTester.Print "light="; LightOn
                 If (LightOn = &HEF Or LightOn = &HCF Or LightOn = 223) Then
                    MPTester.TestResultLab = "PASS "
                    TestResult = "PASS"
                    ContFail = 0
                Else
                 
                  TestResult = "Bin3"
                  MPTester.TestResultLab = "Bin3:LED FAIL "
              
               End If
               
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:Undefine Fail"
          
             ContFail = ContFail + 1
        
               
        End Select
                               
       cardresult = DO_WritePort(card, Channel_P1A, &HFB)
         'Call PowerSet(1500)
        Call PowerSet2(1, "5.0", "0.2", 1, "5.0", "0.2", 1)
                                    
End Sub


Public Sub AU6997A51BLF20TestSub()
'add unload driver function
' use AU6996 new MP and AU6992 RW test

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
'===============================================================3
 
If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
    FileCopy App.Path & "\AU6997FT\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
    Call MsecDelay(5)
End If


NewChipFlag = 0
If OldChipName <> ChipName Then
            FileCopy App.Path & "\AU6997FT\PE_INI\" & ChipName & "\PE.ini", App.Path & "\AU6997FT\PE.ini"
         '   FileCopy App.Path & "\AU6997FT\ROM\" & chipname & "\ROM.Hex", App.Path & "\AU6997FT\ROM.Hex"
         '   FileCopy App.Path & "\AU6997FT\RAM\" & chipname & "\RAM.Bin", App.Path & "\AU6997FT\RAM.Bin"
              FileCopy App.Path & "\AU6997FT\INI\" & ChipName & "\AlcorMP.ini", App.Path & "\AU6997FT\AlcorMP.ini"
            NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName
 

'==============================================================
' when begin RW Test, must clear MP program
'===============================================================
'(1)
winHwnd = FindWindow(vbNullString, AU6997MPCaption1)
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, AU6997MPCaption1)
  Loop While winHwnd <> 0
End If


winHwnd = FindWindow(vbNullString, AU6997MPCaption)
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, AU6997MPCaption)
  Loop While winHwnd <> 0
End If
'(2)
winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
  Loop While winHwnd <> 0
End If
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail
 
 
 '====================================
 '  Fix Card
 '====================================
' GoTo T1
 If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag = True) Then
 
 
   If MPTester.NoMP.Value = 1 Then
        
        If (NewChipFlag = 0) And (MPTester.Check1.Value = 0) Then  ' force condition
           GoTo RW_Test_Label
        End If
    End If
       
    
    If MPTester.ResetMPFailCounter.Value = 1 Then
       ContFail = 0
    End If
    
 '==============================================================
' when begin MP, must close RW program
'===============================================================
   MPFlag = 1
 
    winHwnd = FindWindow(vbNullString, "UFD Test")
    If winHwnd <> 0 Then
      Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "UFD Test")
      Loop While winHwnd <> 0
    End If
 
       '  power on
       cardresult = DO_WritePort(card, Channel_P1A, &HFF)
       ' Call PowerSet(3)   ' close power to disable chip
       Call MsecDelay(0.5)  ' power for load MPDriver
       MPTester.Print "wait for MP Ready"
     '  Call LoadMP_Click_AU6996Reader
        Call LoadMP_Click_AU6997
        OldTimer = Timer
        AlcorMPMessage = 0
        Debug.Print "begin"
        Do
           ' DoEvents
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
            PassTime = Timer - OldTimer
            'Debug.Print AlcorMPMessage
        Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 30 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3
        '====================================================
        If PassTime > 30 Then
          '(1)
           MPTester.Print "MP Ready Fail"
            winHwnd = FindWindow(vbNullString, AU6997MPCaption1)
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, AU6997MPCaption1)
              Loop While winHwnd <> 0
            End If
           '(2)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
            If winHwnd <> 0 Then
              Do
              rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
              Call MsecDelay(0.5)
              winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
              Loop While winHwnd <> 0
            End If
            
        
             MPTester.TestResultLab = "Bin3:MP Ready Fail"
             TestResult = "Bin3"
             MPTester.Print "MP Ready Fail"
     
              
            Exit Sub
        End If
        
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_MP_START Then
         
             
             cardresult = DO_WritePort(card, Channel_P1A, &H0)  ' sel chip
        '      Call PowerSet(500)   ' close power to disable chip
             
 
          
             
            Dim TimerCounter As Integer
            Dim TmpString As String
            
             
            Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName("vid")
            Loop While TmpString = "" And TimerCounter < 150
             
             
             If TmpString = "" Then   ' can not find device after 15 s
             
               TestResult = "Bin2"
               MPTester.TestResultLab = "Bin2:MP UNKNOW Fail when enter MP"
               Exit Sub
             End If
             
             Call MsecDelay(2.5)
              MPTester.Cls
             MPTester.Print " MP Begin....."
             
            ' Call StartMP_Click_AU6996
              Call StartMP_Click_AU6997
             
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
                Loop Until AlcorMPMessage = WM_FT_MP_PASS _
                Or AlcorMPMessage = WM_FT_MP_FAIL _
                Or AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL _
                Or PassTime > 100 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                
                MPTester.Print "MP work time="; PassTime
                 MPTester.MPText.Text = Hex(AlcorMPMessage)
                '================================================
                '  Handle MP work time out error
                '===============================================
                
               ' time out fail
                If PassTime > 100 Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    '(1)
                    winHwnd = FindWindow(vbNullString, AU6997MPCaption)
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, AU6997MPCaption)
                      Loop While winHwnd <> 0
                    End If
                    
                    '(2)
                      winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    Exit Sub
                End If
                
                ' MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    
                    winHwnd = FindWindow(vbNullString, AU6997MPCaption)
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, AU6997MPCaption)
                      Loop While winHwnd <> 0
                    End If
                    
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    
                    Exit Sub
                End If
                
                
                 'unknow fail
                 If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                     MPContFail = MPContFail + 1
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                    MPTester.Print "MP UNKNOW Fail"
                    
                   winHwnd = FindWindow(vbNullString, AU6997MPCaption)
                    If winHwnd <> 0 Then
                     Do
                     rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                     Call MsecDelay(0.5)
                     winHwnd = FindWindow(vbNullString, AU6997MPCaption)
                     Loop While winHwnd <> 0
                     End If
                     
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                     If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                     End If
                     
                     
                     
                    Exit Sub
                End If
                 
                
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_PASS Then
                     MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPTester.Print "MP PASS"
                End If
        End If
   
End If
'=========================================
 '    Close MP program
 '=========================================
winHwnd = FindWindow(vbNullString, AU6997MPCaption)
If winHwnd <> 0 Then
  Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, AU6997MPCaption)
  Loop While winHwnd <> 0
End If
    
 winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
    Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
    Loop While winHwnd <> 0
End If

  Dim pid As Long          ' unload driver
  Dim hProcess As Long
  Dim ExitEvent As Long
 
  pid = Shell(App.Path & "\AU6997FT\loaddrv.exe uninstall_058F6387")
  hProcess = OpenProcess(SYNCHRONIZE + PROCESS_QUERY_INFORMATION + PROCESS_TERMINATE, 0, pid)
  ExitEvent = WaitForSingleObject(hProcess, INFINITE)
  Call CloseHandle(hProcess)
  KillProcess ("LoadDrv.exe")
  
 
                        
 '=========================================
 '    POWER on
 '=========================================
'T1:
RW_Test_Label:
 
 If MPFlag = 1 Then
        ' Call PowerSet(3)
          cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        
         Call MsecDelay(0.5)  ' power of to unload MPDriver

           cardresult = DO_WritePort(card, Channel_P1A, &H0)  'sel socket
         '  Call PowerSet(1500)
     
        
         Call MsecDelay(3.2)
        MPFlag = 0
 Else
          cardresult = DO_WritePort(card, Channel_P1A, &H0)
       '  Call PowerSet(1500)
         
         Call MsecDelay(1.2)
End If
       '  Call LoadRWTest_Click_AU6996Reader
        Call LoadRWTest_Click_AU6997
        
        
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
     '   GoTo T2
       If PassTime > 5 Then
           TestResult = "Bin3"
           MPTester.TestResultLab = "Bin3:RW Ready Fail"
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
       
            Exit Sub
       End If
         
T2:
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        Call StartRWTest_Click_AU6988
        
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
                Or AlcorMPMessage = WM_FT_CHECK_CERBGPO_FAIL _
                  Or AlcorMPMessage = WM_FT_PARAM_FAIL _
                  Or AlcorMPMessage = WM_FT_READER_FAIL _
                  Or AlcorMPMessage = WM_FT_BUSWIDTH_FAIL _
                  Or AlcorMPMessage = WM_FT_BUSCLK_FAIL _
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
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
            
              cardresult = DO_WritePort(card, Channel_P1A, &HFF)  ' power off
            
          '   Call PowerSet(1500)
        
            
       
            Exit Sub
        End If
        
        
     
               
        Select Case AlcorMPMessage
        
        Case WM_FT_RW_UNKNOW_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:UnKnow Fail"
          
             ContFail = ContFail + 1
        
        Case WM_FT_RW_SPEED_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:SPEED Error "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_RW_FAIL, WM_FT_PARAM_FAIL, WM_FT_READER_FAIL, WM_FT_BUSWIDTH_FAIL, WM_FT_BUSCLK_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:RW FAIL "
             ContFail = ContFail + 1
             
        Case WM_FT_CHECK_CERBGPO_FAIL
        
             TestResult = "Bin3"
             MPTester.TestResultLab = "Bin3:GPO/RB FAIL "
             ContFail = ContFail + 1
             
        Case WM_FT_RW_ROM_FAIL
              TestResult = "Bin4"
              MPTester.TestResultLab = "Bin4:ROM FAIL "
              ContFail = ContFail + 1
              
        Case WM_FT_RW_RAM_FAIL
              TestResult = "Bin5"
              MPTester.TestResultLab = "Bin5:RAM FAIL "
               ContFail = ContFail + 1
        Case WM_FT_RW_RW_PASS
        
               
               For LedCount = 1 To 20
               Call MsecDelay(0.1)
               cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
                If LightOn = &HFE Then
               
                 Exit For
               End If
               Next LedCount
                 
                  MPTester.Print "light="; LightOn
                 If LightOn = &HFE Then
                    MPTester.TestResultLab = "PASS "
                    TestResult = "PASS"
                    ContFail = 0
                Else
                 
                  TestResult = "Bin3"
                  MPTester.TestResultLab = "Bin3:LED FAIL "
              
               End If
               
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:Undefine Fail"
          
             ContFail = ContFail + 1
        
               
        End Select
                               
       cardresult = DO_WritePort(card, Channel_P1A, &HFB)
       '  Call PowerSet(1500)

End Sub

Public Sub AU6996A51ILF2ATestSub()
'add unload driver function
' use AU6996 new MP and AU6992 RW test

 If PCI7248InitFinish = 0 Then
       PCI7248Exist
 End If
 Dim OldTimer
 Dim PassTime
 Dim rt2
 Dim LightOn
 Dim mMsg As MSG
 Dim LedCount As Byte
 'dim remp_flag As Byte
 
   MPTester.TestResultLab = ""
'===============================================================
' Fail location initial
'===============================================================3
 
If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
    FileCopy App.Path & "\AlcorMP_698x2\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
    Call MsecDelay(5)
End If


NewChipFlag = 0
If OldChipName <> ChipName Then
            FileCopy App.Path & "\AlcorMP_6996Flash\New_INI\" & ChipName & "\ROM\ROM.Hex", App.Path & "\AlcorMP_6996Flash\ROM.Hex"
            FileCopy App.Path & "\AlcorMP_6996Flash\New_INI\" & ChipName & "\RAM\RAM.Bin", App.Path & "\AlcorMP_6996Flash\RAM.Bin"
            FileCopy App.Path & "\AlcorMP_6996Flash\New_INI\" & ChipName & "\AlcorMP.ini", App.Path & "\AlcorMP_6996Flash\AlcorMP.ini"
            FileCopy App.Path & "\AlcorMP_6996Flash\New_INI\" & ChipName & "\PE.bin", App.Path & "\AlcorMP_6996Flash\PE.bin"
            NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName
 

'==============================================================
' when begin RW Test, must clear MP program
'===============================================================
'(1)
winHwnd = FindWindow(vbNullString, AU6996MPCaption1)
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, AU6996MPCaption1)
  Loop While winHwnd <> 0
End If


winHwnd = FindWindow(vbNullString, AU6996MPCaption)
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, AU6996MPCaption)
  Loop While winHwnd <> 0
End If
'(2)
winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
  Loop While winHwnd <> 0
End If
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail
 
 
 '====================================
 '  Fix Card
 '====================================
' GoTo T1
 If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag = True) Then
 
 
   If MPTester.NoMP.Value = 1 Then
        
        If (NewChipFlag = 0) And (MPTester.Check1.Value = 0) Then  ' force condition
           GoTo RW_Test_Label
        End If
    End If
       
    
    If MPTester.ResetMPFailCounter.Value = 1 Then
       ContFail = 0
    End If
    
 '==============================================================
' when begin MP, must close RW program
'===============================================================
   MPFlag = 1
 
    winHwnd = FindWindow(vbNullString, "UFD Test")
    If winHwnd <> 0 Then
      Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "UFD Test")
      Loop While winHwnd <> 0
    End If
 
       '  power on
       cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        Call PowerSet(3)   ' close power to disable chip
       Call MsecDelay(0.5)  ' power for load MPDriver
       MPTester.Print "wait for MP Ready"
       Call LoadNewMP_Click_AU6996Flash
 
        OldTimer = Timer
        AlcorMPMessage = 0
        Debug.Print "begin"
        Do
           ' DoEvents
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
            PassTime = Timer - OldTimer
            'Debug.Print AlcorMPMessage
        Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 30 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3
        '====================================================
        If PassTime > 30 Then
          '(1)
           MPTester.Print "MP Ready Fail"
            winHwnd = FindWindow(vbNullString, AU6996MPCaption1)
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, AU6996MPCaption1)
              Loop While winHwnd <> 0
            End If
           '(2)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
            If winHwnd <> 0 Then
              Do
              rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
              Call MsecDelay(0.5)
              winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
              Loop While winHwnd <> 0
            End If
            
        
             MPTester.TestResultLab = "Bin3:MP Ready Fail"
             TestResult = "Bin3"
             MPTester.Print "MP Ready Fail"
     
              
            Exit Sub
        End If
        
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_MP_START Then
         
             
             cardresult = DO_WritePort(card, Channel_P1A, &HFB)  ' sel chip
              Call PowerSet(500)   ' close power to disable chip
             
 
          
             
            Dim TimerCounter As Integer
            Dim TmpString As String
            
             
            Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName("vid")
            Loop While TmpString = "" And TimerCounter < 150
             
            Call MsecDelay(0.3)
            
             If TmpString = "" Then   ' can not find device after 15 s
             
               TestResult = "Bin2"
               MPTester.TestResultLab = "Bin2:MP UNKNOW Fail when enter MP"
               Exit Sub
             End If
             
             Call MsecDelay(2.5)
               
             MPTester.Print " MP Begin....."
             
             Call StartMP_Click_AU6996
   
             
              OldTimer = Timer
              AlcorMPMessage = 0
              ReMP_Flag = 0
              
                Do
                   ' DoEvents
                    If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                        AlcorMPMessage = mMsg.message
                        TranslateMessage mMsg
                        DispatchMessage mMsg
                        
                        If (AlcorMPMessage = WM_FT_MP_FAIL) And (ReMP_Flag = 0) Then
                            'ReMP_Flag = 1
                            AlcorMPMessage = 1
                            cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'close power
                            Call MsecDelay(0.3)
                            cardresult = DO_WritePort(card, Channel_P1A, &HFB)  'open power
                            Call MsecDelay(2.2)
                            Call RefreshMP_Click_AU6996
                            Call MsecDelay(0.5)
                            Call StartMP_Click_AU6996
                            
                            ReMP_Counter = ReMP_Counter + 1
                            If ReMP_Counter >= ReMP_Limit Then
                                ReMP_Flag = 1
                                ReMP_Counter = 0
                            End If
                        End If
                    End If
                    PassTime = Timer - OldTimer
                Loop Until AlcorMPMessage = WM_FT_MP_PASS _
                Or AlcorMPMessage = WM_FT_MP_FAIL _
                Or AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL _
                Or PassTime > 80 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                
                If AlcorMPMessage = WM_FT_MP_PASS Then
                    ReMP_Counter = 0
                End If
                
                MPTester.Print "MP work time="; PassTime
                 MPTester.MPText.Text = Hex(AlcorMPMessage)
                '================================================
                '  Handle MP work time out error
                '===============================================
                
               ' time out fail
                If PassTime > 65 Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    '(1)
                    winHwnd = FindWindow(vbNullString, AU6996MPCaption)
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, AU6996MPCaption)
                      Loop While winHwnd <> 0
                    End If
                    
                    '(2)
                      winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    Exit Sub
                End If
                
                ' MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    
                    winHwnd = FindWindow(vbNullString, AU6996MPCaption)
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, AU6996MPCaption)
                      Loop While winHwnd <> 0
                    End If
                    
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    
                    Exit Sub
                End If
                
                
                 'unknow fail
                 If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                     MPContFail = MPContFail + 1
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                    MPTester.Print "MP UNKNOW Fail"
                    
                   winHwnd = FindWindow(vbNullString, AU6996MPCaption)
                    If winHwnd <> 0 Then
                     Do
                     rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                     Call MsecDelay(0.5)
                     winHwnd = FindWindow(vbNullString, AU6996MPCaption)
                     Loop While winHwnd <> 0
                     End If
                     
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                     If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                     End If
                     
                     
                     
                    Exit Sub
                End If
                 
                
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_PASS Then
                     MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPTester.Print "MP PASS"
                End If
        End If
   
End If
'=========================================
 '    Close MP program
 '=========================================
winHwnd = FindWindow(vbNullString, AU6996MPCaption)
If winHwnd <> 0 Then
  Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, AU6996MPCaption)
  Loop While winHwnd <> 0

    Call MsecDelay(0.2)
    KillProcess ("AlcorMP.exe")

End If


 winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
    Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
    Loop While winHwnd <> 0
End If

  Dim pid As Long          ' unload driver
  Dim hProcess As Long
  Dim ExitEvent As Long
 
  pid = Shell(App.Path & "\AlcorMP_6996Flash\loaddrv.exe uninstall_058F6387")
  hProcess = OpenProcess(SYNCHRONIZE + PROCESS_QUERY_INFORMATION + PROCESS_TERMINATE, 0, pid)
  ExitEvent = WaitForSingleObject(hProcess, INFINITE)
  Call CloseHandle(hProcess)
  KillProcess ("LoadDrv.exe")
 
 
                        
 '=========================================
 '    POWER on
 '=========================================
'T1:
RW_Test_Label:
 
 If MPFlag = 1 Then
         Call PowerSet(3)
          cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        
         Call MsecDelay(0.5)  ' power of to unload MPDriver

           cardresult = DO_WritePort(card, Channel_P1A, &HFB)  'sel socket
           Call PowerSet(1500)
     
        
         Call MsecDelay(2.2)
        MPFlag = 0
 Else
          cardresult = DO_WritePort(card, Channel_P1A, &HFB)
         Call PowerSet(1500)
         
         Call MsecDelay(2.2)
End If
         Call LoadRWTest_Click_AU6996Flash

        
        
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
     '   GoTo T2
       If PassTime > 5 Then
           TestResult = "Bin3"
           MPTester.TestResultLab = "Bin3:RW Ready Fail"
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
       
            Exit Sub
       End If
         
T2:
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        Call StartRWTest_Click_AU6988
        
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
                Or AlcorMPMessage = WM_FT_CHECK_CERBGPO_FAIL _
              Or PassTime > 5 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
    
          MPTester.Print "RW work Time="; PassTime
          MPTester.MPText.Text = Hex(AlcorMPMessage)
        
        
        '===========================================================
        '  RW Time Out Fail
        '===========================================================
        
        If (PassTime > 5) Or ((FailCloseAP) And (AlcorMPMessage <> WM_FT_RW_RW_PASS)) Then
            Close_FT_AP ("UFD Test")
            
            If (PassTime > 5) Then
                TestResult = "Bin3"
                MPTester.TestResultLab = "Bin3:RW Time Out Fail"
                cardresult = DO_WritePort(card, Channel_P1A, &HFB)  ' power off
                Exit Sub
            End If
        
        End If
        
               
        Select Case AlcorMPMessage
        
        Case WM_FT_RW_UNKNOW_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:UnKnow Fail"
          
             ContFail = ContFail + 1
        
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
             
        Case WM_FT_RW_ROM_FAIL
              TestResult = "Bin4"
              MPTester.TestResultLab = "Bin4:ROM FAIL "
              ContFail = ContFail + 1
              
        Case WM_FT_RW_RAM_FAIL
              TestResult = "Bin5"
              MPTester.TestResultLab = "Bin5:RAM FAIL "
               ContFail = ContFail + 1
        Case WM_FT_RW_RW_PASS
        
               
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
              
               End If
               
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:Undefine Fail"
          
             ContFail = ContFail + 1
        
               
        End Select
                               
       cardresult = DO_WritePort(card, Channel_P1A, &HFB)
         Call PowerSet(1500)

End Sub

Public Sub AU6996Flash_DualTestSub()
'add unload driver function
' use AU6996 new MP and AU6992 RW test

    Lon = False
    Loff = False
    LightSituation = 255

 If PCI7248InitFinish = 0 Then
       PCI7248Exist
 End If
 Dim OldTimer
 Dim PassTime
 Dim rt2
 Dim LightOn
 Dim mMsg As MSG
 Dim LedCount As Byte
 'dim remp_flag As Byte
 
   MPTester.TestResultLab = ""
'===============================================================
' Fail location initial
'===============================================================3
 
If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
    FileCopy App.Path & "\AlcorMP_698x2\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
    Call MsecDelay(5)
End If


NewChipFlag = 0
If OldChipName <> ChipName Then
            FileCopy App.Path & "\AlcorMP_6996Flash\New_INI\" & ChipName & "\ROM\ROM.Hex", App.Path & "\AlcorMP_6996Flash\ROM.Hex"
            FileCopy App.Path & "\AlcorMP_6996Flash\New_INI\" & ChipName & "\RAM\RAM.Bin", App.Path & "\AlcorMP_6996Flash\RAM.Bin"
            FileCopy App.Path & "\AlcorMP_6996Flash\New_INI\" & ChipName & "\AlcorMP.ini", App.Path & "\AlcorMP_6996Flash\AlcorMP.ini"
            FileCopy App.Path & "\AlcorMP_6996Flash\New_INI\" & ChipName & "\PE.bin", App.Path & "\AlcorMP_6996Flash\PE.bin"
            NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName
 

'==============================================================
' when begin RW Test, must clear MP program
'===============================================================
'(1)
winHwnd = FindWindow(vbNullString, AU6996MPCaption1)
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, AU6996MPCaption1)
  Loop While winHwnd <> 0
End If


winHwnd = FindWindow(vbNullString, AU6996MPCaption)
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, AU6996MPCaption)
  Loop While winHwnd <> 0
End If
'(2)
winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
  Loop While winHwnd <> 0
End If
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail
 
 
 '====================================
 '  Fix Card
 '====================================
' GoTo T1
 If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag = True) Then
 
 
   If MPTester.NoMP.Value = 1 Then
        
        If (NewChipFlag = 0) And (MPTester.Check1.Value = 0) Then  ' force condition
           GoTo RW_Test_Label
        End If
    End If
       
    
    If MPTester.ResetMPFailCounter.Value = 1 Then
       ContFail = 0
    End If
    
 '==============================================================
' when begin MP, must close RW program
'===============================================================
   MPFlag = 1
 
    winHwnd = FindWindow(vbNullString, "UFD Test")
    If winHwnd <> 0 Then
      Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "UFD Test")
      Loop While winHwnd <> 0
    End If
 
       '  power on
       cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        Call PowerSet(3)   ' close power to disable chip
       Call MsecDelay(0.5)  ' power for load MPDriver
       MPTester.Print "wait for MP Ready"
       Call LoadNewMP_Click_AU6996Flash
 
        OldTimer = Timer
        AlcorMPMessage = 0
        Debug.Print "begin"
        Do
           ' DoEvents
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
            PassTime = Timer - OldTimer
            'Debug.Print AlcorMPMessage
        Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 30 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3
        '====================================================
        If PassTime > 30 Then
          '(1)
           MPTester.Print "MP Ready Fail"
            winHwnd = FindWindow(vbNullString, AU6996MPCaption1)
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, AU6996MPCaption1)
              Loop While winHwnd <> 0
            End If
           '(2)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
            If winHwnd <> 0 Then
              Do
              rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
              Call MsecDelay(0.5)
              winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
              Loop While winHwnd <> 0
            End If
            
        
             MPTester.TestResultLab = "Bin3:MP Ready Fail"
             TestResult = "Bin3"
             MPTester.Print "MP Ready Fail"
     
              
            Exit Sub
        End If
        
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_MP_START Then
         
             
             cardresult = DO_WritePort(card, Channel_P1A, &HFD)  ' sel chip
              Call PowerSet(500)   ' close power to disable chip
             
 
          
             
            Dim TimerCounter As Integer
            Dim TmpString As String
            
             
            Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName("vid")
            Loop While TmpString = "" And TimerCounter < 150
             
            Call MsecDelay(0.3)
            
             If TmpString = "" Then   ' can not find device after 15 s
             
               TestResult = "Bin2"
               MPTester.TestResultLab = "Bin2:MP UNKNOW Fail when enter MP"
               Exit Sub
             End If
             
             Call MsecDelay(2.5)
               
             MPTester.Print " MP Begin....."
             
             Call StartMP_Click_AU6996
   
             
              OldTimer = Timer
              AlcorMPMessage = 0
              ReMP_Flag = 0
              
                Do
                   ' DoEvents
                    If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                        AlcorMPMessage = mMsg.message
                        TranslateMessage mMsg
                        DispatchMessage mMsg
                        
                        If (AlcorMPMessage = WM_FT_MP_FAIL) And (ReMP_Flag = 0) Then
                            'ReMP_Flag = 1
                            AlcorMPMessage = 1
                            cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'close power
                            Call MsecDelay(0.3)
                            cardresult = DO_WritePort(card, Channel_P1A, &HFD)  'open power
                            Call MsecDelay(2.2)
                            Call RefreshMP_Click_AU6996
                            Call MsecDelay(0.5)
                            Call StartMP_Click_AU6996
                            
                            ReMP_Counter = ReMP_Counter + 1
                            If ReMP_Counter >= ReMP_Limit Then
                                ReMP_Flag = 1
                                ReMP_Counter = 0
                            End If
                        End If
                    End If
                    PassTime = Timer - OldTimer
                Loop Until AlcorMPMessage = WM_FT_MP_PASS _
                Or AlcorMPMessage = WM_FT_MP_FAIL _
                Or AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL _
                Or PassTime > 80 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                
                If AlcorMPMessage = WM_FT_MP_PASS Then
                    ReMP_Counter = 0
                End If
                
                MPTester.Print "MP work time="; PassTime
                 MPTester.MPText.Text = Hex(AlcorMPMessage)
                '================================================
                '  Handle MP work time out error
                '===============================================
                
               ' time out fail
                If PassTime > 65 Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    '(1)
                    winHwnd = FindWindow(vbNullString, AU6996MPCaption)
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, AU6996MPCaption)
                      Loop While winHwnd <> 0
                    End If
                    
                    '(2)
                      winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    Exit Sub
                End If
                
                ' MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    
                    winHwnd = FindWindow(vbNullString, AU6996MPCaption)
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, AU6996MPCaption)
                      Loop While winHwnd <> 0
                    End If
                    
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    
                    Exit Sub
                End If
                
                
                 'unknow fail
                 If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                     MPContFail = MPContFail + 1
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                    MPTester.Print "MP UNKNOW Fail"
                    
                   winHwnd = FindWindow(vbNullString, AU6996MPCaption)
                    If winHwnd <> 0 Then
                     Do
                     rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                     Call MsecDelay(0.5)
                     winHwnd = FindWindow(vbNullString, AU6996MPCaption)
                     Loop While winHwnd <> 0
                     End If
                     
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                     If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                     End If
                     
                     
                     
                    Exit Sub
                End If
                 
                
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_PASS Then
                     MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPTester.Print "MP PASS"
                End If
        End If
   
End If
'=========================================
 '    Close MP program
 '=========================================
winHwnd = FindWindow(vbNullString, AU6996MPCaption)
If winHwnd <> 0 Then
  Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, AU6996MPCaption)
  Loop While winHwnd <> 0

    Call MsecDelay(0.2)
    KillProcess ("AlcorMP.exe")

End If


 winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
    Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
    Loop While winHwnd <> 0
End If

  Dim pid As Long          ' unload driver
  Dim hProcess As Long
  Dim ExitEvent As Long
 
  pid = Shell(App.Path & "\AlcorMP_6996Flash\loaddrv.exe uninstall_058F6387")
  hProcess = OpenProcess(SYNCHRONIZE + PROCESS_QUERY_INFORMATION + PROCESS_TERMINATE, 0, pid)
  ExitEvent = WaitForSingleObject(hProcess, INFINITE)
  Call CloseHandle(hProcess)
  KillProcess ("LoadDrv.exe")
 
 
                        
 '=========================================
 '    POWER on
 '=========================================
'T1:
RW_Test_Label:
 
 If MPFlag = 1 Then
         Call PowerSet(3)
          cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        
         Call MsecDelay(0.5)  ' power of to unload MPDriver

           cardresult = DO_WritePort(card, Channel_P1A, &HFA)  'sel socket
           Call PowerSet(1500)
     
        
         Call MsecDelay(2.2)
        MPFlag = 0
 Else
          cardresult = DO_WritePort(card, Channel_P1A, &HFA)
         Call PowerSet(1500)
         
         Call MsecDelay(2.2)
End If
         Call LoadRWTest_Click_AU6996Flash

        ' 20131129 add for led check
        For LedCount = 1 To 10
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
     '   GoTo T2
       If PassTime > 5 Then
           TestResult = "Bin3"
           MPTester.TestResultLab = "Bin3:RW Ready Fail"
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
       
            Exit Sub
       End If
         
T2:
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        Call StartRWTest_Click_AU6988
        
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
                Or AlcorMPMessage = WM_FT_CHECK_CERBGPO_FAIL _
              Or PassTime > 5 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
    
          MPTester.Print "RW work Time="; PassTime
          MPTester.MPText.Text = Hex(AlcorMPMessage)
        
        
        '===========================================================
        '  RW Time Out Fail
        '===========================================================
        
        If (PassTime > 5) Or ((FailCloseAP) And (AlcorMPMessage <> WM_FT_RW_RW_PASS)) Then
            Close_FT_AP ("UFD Test")
            
            If (PassTime > 5) Then
                TestResult = "Bin3"
                MPTester.TestResultLab = "Bin3:RW Time Out Fail"
                cardresult = DO_WritePort(card, Channel_P1A, &HFA)  ' power off
                Exit Sub
            End If
        
        End If
        
               
        Select Case AlcorMPMessage
        
        Case WM_FT_RW_UNKNOW_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:UnKnow Fail"
          
             ContFail = ContFail + 1
        
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
             
        Case WM_FT_RW_ROM_FAIL
              TestResult = "Bin4"
              MPTester.TestResultLab = "Bin4:ROM FAIL "
              ContFail = ContFail + 1
              
        Case WM_FT_RW_RAM_FAIL
              TestResult = "Bin5"
              MPTester.TestResultLab = "Bin5:RAM FAIL "
               ContFail = ContFail + 1
        Case WM_FT_RW_RW_PASS
        
               
               For LedCount = 1 To 20
               Call MsecDelay(0.1)
               cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
                If LightOn = &HEF Or LightOn = &HCF Or LightOn = 223 Then
               
                 Exit For
               End If
               Next LedCount
                 
                 ' 20131129 add check led off
                 MPTester.Print "light="; LightOn
                 If (LightOn = &HEF Or LightOn = &HCF Or LightOn = 223) And (Loff = True) Then
                    MPTester.TestResultLab = "PASS "
                    TestResult = "PASS"
                    ContFail = 0
                Else
                 
                  TestResult = "Bin3"
                  MPTester.TestResultLab = "Bin3:LED FAIL "
              
               End If
               
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:Undefine Fail"
          
             ContFail = ContFail + 1
        
               
        End Select
                               
       cardresult = DO_WritePort(card, Channel_P1A, &HFA)
         Call PowerSet(1500)

End Sub

Public Sub AU6992A52HLF20TestSub()
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
 
If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
    FileCopy App.Path & "\AlcorMP_698x2\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
    Call MsecDelay(5)
End If


NewChipFlag = 0
If OldChipName <> ChipName Then
            FileCopy App.Path & "\AlcorMP_698x2\ROM\" & ChipName & "\ROM.Hex", App.Path & "\AlcorMP_698x2\ROM.Hex"
            FileCopy App.Path & "\AlcorMP_698x2\RAM\" & ChipName & "\RAM.Bin", App.Path & "\AlcorMP_698x2\RAM.Bin"
            FileCopy App.Path & "\AlcorMP_698x2\INI\" & ChipName & "\AlcorMP.ini", App.Path & "\AlcorMP_698x2\AlcorMP.ini"
            NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName
 

'==============================================================
' when begin RW Test, must clear MP program
'===============================================================


'(1)  /// for Auto mode
winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
  Loop While winHwnd <> 0
End If

'(1)
winHwnd = FindWindow(vbNullString, AU6992MPCaption)
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, AU6992MPCaption)
  Loop While winHwnd <> 0
End If
'(2)
winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
  Loop While winHwnd <> 0
End If
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail
 
 
 '====================================
 '  Fix Card
 '====================================
' GoTo T1
 If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag = True) Then
 
 
   If MPTester.NoMP.Value = 1 Then
        
        If (NewChipFlag = 0) And (MPTester.Check1.Value = 0) Then  ' force condition
           GoTo RW_Test_Label
        End If
    End If
       
    
    If MPTester.ResetMPFailCounter.Value = 1 Then
       ContFail = 0
    End If
    
 '==============================================================
' when begin MP, must close RW program
'===============================================================
   MPFlag = 1
 
    winHwnd = FindWindow(vbNullString, "UFD Test")
    If winHwnd <> 0 Then
      Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "UFD Test")
      Loop While winHwnd <> 0
    End If
 
       '  power on
       cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        Call PowerSet(3)   ' close power to disable chip
       Call MsecDelay(0.5)  ' power for load MPDriver
       MPTester.Print "wait for MP Ready"
       Call LoadMP_Click_AU6992
 
        OldTimer = Timer
        AlcorMPMessage = 0
        Debug.Print "begin"
        Do
           ' DoEvents
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
            PassTime = Timer - OldTimer
            'Debug.Print AlcorMPMessage
        Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 30 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3
        '====================================================
        If PassTime > 30 Then
          '(1)
           MPTester.Print "MP Ready Fail"
            winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
              Loop While winHwnd <> 0
            End If
           '(2)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
            If winHwnd <> 0 Then
              Do
              rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
              Call MsecDelay(0.5)
              winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
              Loop While winHwnd <> 0
            End If
            
        
             MPTester.TestResultLab = "Bin3:MP Ready Fail"
             TestResult = "Bin3"
             MPTester.Print "MP Ready Fail"
     
              
            Exit Sub
        End If
        
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_MP_START Then
         
             
             cardresult = DO_WritePort(card, Channel_P1A, &HFB)  ' sel chip
              Call PowerSet(500)   ' close power to disable chip
             
 
          
             
            Dim TimerCounter As Integer
            Dim TmpString As String
            
             
            Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName("vid")
            Loop While TmpString = "" And TimerCounter < 150
             
             
             If TmpString = "" Then   ' can not find device after 15 s
             
               TestResult = "Bin2"
               MPTester.TestResultLab = "Bin2:MP UNKNOW Fail when enter MP"
               Exit Sub
             End If
             
             Call MsecDelay(2.5)
               
             MPTester.Print " MP Begin....."
             
             Call StartMP_Click_AU6992
   
             
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
                
               ' time out fail
                If PassTime > 65 Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    '(1)
                    winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                      Loop While winHwnd <> 0
                    End If
                    
                    '(2)
                      winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    Exit Sub
                End If
                
                ' MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    
                    winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                      Loop While winHwnd <> 0
                    End If
                    
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    
                    Exit Sub
                End If
                
                
                 'unknow fail
                 If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                     MPContFail = MPContFail + 1
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                    MPTester.Print "MP UNKNOW Fail"
                    
                   winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                    If winHwnd <> 0 Then
                     Do
                     rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                     Call MsecDelay(0.5)
                     winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                     Loop While winHwnd <> 0
                     End If
                     
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                     If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                     End If
                     
                     
                     
                    Exit Sub
                End If
                 
                
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_PASS Then
                     MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPTester.Print "MP PASS"
                End If
        End If
   
End If
'=========================================
 '    Close MP program
 '=========================================
winHwnd = FindWindow(vbNullString, AU6992MPCaption)
If winHwnd <> 0 Then
  Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, AU6992MPCaption)
  Loop While winHwnd <> 0
End If
    
 winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
    Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
    Loop While winHwnd <> 0
End If

  Dim pid As Long          ' unload driver
  Dim hProcess As Long
  Dim ExitEvent As Long
 
  pid = Shell(App.Path & "\AlcorMP_698x2\loaddrv.exe uninstall_058F6387")
  hProcess = OpenProcess(SYNCHRONIZE + PROCESS_QUERY_INFORMATION + PROCESS_TERMINATE, 0, pid)
  ExitEvent = WaitForSingleObject(hProcess, INFINITE)
  Call CloseHandle(hProcess)
  KillProcess ("LoadDrv.exe")
 
 
                        
 '=========================================
 '    POWER on
 '=========================================
'T1:
RW_Test_Label:
 
 If MPFlag = 1 Then
         Call PowerSet(3)
          cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        
         Call MsecDelay(0.5)  ' power of to unload MPDriver

           cardresult = DO_WritePort(card, Channel_P1A, &HFB)  'sel socket
           Call PowerSet(1500)
     
        
         Call MsecDelay(1.2)
        MPFlag = 0
 Else
          cardresult = DO_WritePort(card, Channel_P1A, &HFB)
         Call PowerSet(1500)
         
         Call MsecDelay(1.2)
End If
         Call LoadRWTest_Click_AU6992

        
        
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
     '   GoTo T2
       If PassTime > 5 Then
           TestResult = "Bin3"
           MPTester.TestResultLab = "Bin3:RW Ready Fail"
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
       
            Exit Sub
       End If
         
T2:
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        Call StartRWTest_Click_AU6988
        
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
                Or AlcorMPMessage = WM_FT_CHECK_CERBGPO_FAIL _
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
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
            
              cardresult = DO_WritePort(card, Channel_P1A, &HFB)  ' power off
            
             Call PowerSet(1500)
        
            
       
            Exit Sub
        End If
        
        
     
               
        Select Case AlcorMPMessage
        
        Case WM_FT_RW_UNKNOW_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:UnKnow Fail"
          
             ContFail = ContFail + 1
        
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
             
        Case WM_FT_RW_ROM_FAIL
              TestResult = "Bin4"
              MPTester.TestResultLab = "Bin4:ROM FAIL "
              ContFail = ContFail + 1
              
        Case WM_FT_RW_RAM_FAIL
              TestResult = "Bin5"
              MPTester.TestResultLab = "Bin5:RAM FAIL "
               ContFail = ContFail + 1
        Case WM_FT_RW_RW_PASS
        
               
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
              
               End If
               
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:Undefine Fail"
          
             ContFail = ContFail + 1
        
               
        End Select
                               
       cardresult = DO_WritePort(card, Channel_P1A, &HFB)
         Call PowerSet(1500)
                            
End Sub

Public Sub AU6992A51HLF2ATestSub()
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
 'dim remp_flag As Byte
   
   MPTester.TestResultLab = ""
'===============================================================
' Fail location initial
'===============================================================
 
If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
    FileCopy App.Path & "\AlcorMP_698x2\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
    Call MsecDelay(5)
End If


NewChipFlag = 0
If OldChipName <> ChipName Then
            FileCopy App.Path & "\AlcorMP_698x2\New_INI\" & ChipName & "\ROM\ROM.Hex", App.Path & "\AlcorMP_698x2\ROM.Hex"
            FileCopy App.Path & "\AlcorMP_698x2\New_INI\" & ChipName & "\RAM\\RAM.Bin", App.Path & "\AlcorMP_698x2\RAM.Bin"
            FileCopy App.Path & "\AlcorMP_698x2\New_INI\" & ChipName & "\AlcorMP.ini", App.Path & "\AlcorMP_698x2\AlcorMP.ini"
            FileCopy App.Path & "\AlcorMP_698x2\New_INI\" & ChipName & "\PE.bin", App.Path & "\AlcorMP_698x2\PE.bin"
            NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName
 

'==============================================================
' when begin RW Test, must clear MP program
'===============================================================


'(1)  /// for Auto mode
winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
  Loop While winHwnd <> 0
End If

'(1)
winHwnd = FindWindow(vbNullString, AU6992MPCaption)
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, AU6992MPCaption)
  Loop While winHwnd <> 0
End If
'(2)
winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
  Loop While winHwnd <> 0
End If
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail
 
 
 '====================================
 '  Fix Card
 '====================================
' GoTo T1
 If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag = True) Then
 
 
   If MPTester.NoMP.Value = 1 Then
        
        If (NewChipFlag = 0) And (MPTester.Check1.Value = 0) Then  ' force condition
           GoTo RW_Test_Label
        End If
    End If
       
    
    If MPTester.ResetMPFailCounter.Value = 1 Then
       ContFail = 0
    End If
    
 '==============================================================
' when begin MP, must close RW program
'===============================================================
   MPFlag = 1
 
    winHwnd = FindWindow(vbNullString, "UFD Test")
    If winHwnd <> 0 Then
      Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "UFD Test")
      Loop While winHwnd <> 0
    End If
 
       '  power on
       cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        Call PowerSet(3)   ' close power to disable chip
       Call MsecDelay(0.5)  ' power for load MPDriver
       MPTester.Print "wait for MP Ready"
       Call LoadMP_Click_AU6992
 
        OldTimer = Timer
        AlcorMPMessage = 0
        Debug.Print "begin"
        Do
           ' DoEvents
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
            PassTime = Timer - OldTimer
            'Debug.Print AlcorMPMessage
        Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 30 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3
        '====================================================
        If PassTime > 30 Then
          '(1)
           MPTester.Print "MP Ready Fail"
            winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
              Loop While winHwnd <> 0
            End If
           '(2)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
            If winHwnd <> 0 Then
              Do
              rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
              Call MsecDelay(0.5)
              winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
              Loop While winHwnd <> 0
            End If
            
        
             MPTester.TestResultLab = "Bin3:MP Ready Fail"
             TestResult = "Bin3"
             MPTester.Print "MP Ready Fail"
     
              
            Exit Sub
        End If
        
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_MP_START Then
         
             
             cardresult = DO_WritePort(card, Channel_P1A, &HFB)  ' sel chip
              Call PowerSet(500)   ' close power to disable chip
             
 
          
             
            Dim TimerCounter As Integer
            Dim TmpString As String
            
             
            Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName("vid")
            Loop While TmpString = "" And TimerCounter < 150
             
            Call MsecDelay(0.3)
            
             If TmpString = "" Then   ' can not find device after 15 s
             
               TestResult = "Bin2"
               MPTester.TestResultLab = "Bin2:MP UNKNOW Fail when enter MP"
               Exit Sub
             End If
             
             Call MsecDelay(2.5)
               
             MPTester.Print " MP Begin....."
             
             Call StartMP_Click_AU6992
   
             
              OldTimer = Timer
              AlcorMPMessage = 0
              ReMP_Flag = 0
              
                Do
                   ' DoEvents
                    If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                        AlcorMPMessage = mMsg.message
                        TranslateMessage mMsg
                        DispatchMessage mMsg
                        
                        If (AlcorMPMessage = WM_FT_MP_FAIL) And (ReMP_Flag = 0) Then
                            'ReMP_Flag = 1
                            AlcorMPMessage = 1
                            cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'close power
                            Call MsecDelay(0.3)
                            cardresult = DO_WritePort(card, Channel_P1A, &HFB)  'open power
                            Call MsecDelay(2.2)
                            Call RefreshMP_Click_AU6992
                            Call MsecDelay(0.5)
                            Call StartMP_Click_AU6992
                            
                            ReMP_Counter = ReMP_Counter + 1
                            If ReMP_Counter >= ReMP_Limit Then
                                ReMP_Flag = 1
                                ReMP_Counter = 0
                            End If
                        End If
                    
                    End If
                    PassTime = Timer - OldTimer
                Loop Until AlcorMPMessage = WM_FT_MP_PASS _
                Or AlcorMPMessage = WM_FT_MP_FAIL _
                Or AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL _
                Or PassTime > 65 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                
                If AlcorMPMessage = WM_FT_MP_PASS Then
                    ReMP_Counter = 0
                End If
                
                MPTester.Print "MP work time="; PassTime
                 MPTester.MPText.Text = Hex(AlcorMPMessage)
                '================================================
                '  Handle MP work time out error
                '===============================================
                
               ' time out fail
                If PassTime > 65 Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    '(1)
                    winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                      Loop While winHwnd <> 0
                    End If
                    
                    '(2)
                      winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    Exit Sub
                End If
                
                ' MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    
                    winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                      Loop While winHwnd <> 0
                    End If
                    
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    
                    Exit Sub
                End If
                
                
                 'unknow fail
                 If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                     MPContFail = MPContFail + 1
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                    MPTester.Print "MP UNKNOW Fail"
                    
                   winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                    If winHwnd <> 0 Then
                     Do
                     rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                     Call MsecDelay(0.5)
                     winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                     Loop While winHwnd <> 0
                     End If
                     
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                     If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                     End If
                     
                     
                     
                    Exit Sub
                End If
                 
                
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_PASS Then
                     MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPTester.Print "MP PASS"
                End If
        End If
   
End If
'=========================================
 '    Close MP program
 '=========================================
winHwnd = FindWindow(vbNullString, AU6992MPCaption)
If winHwnd <> 0 Then
  Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, AU6992MPCaption)
  Loop While winHwnd <> 0
    Call MsecDelay(0.2)
    KillProcess ("AlcorMP.exe")

End If
    
 winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
    Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
    Loop While winHwnd <> 0
End If

  Dim pid As Long          ' unload driver
  Dim hProcess As Long
  Dim ExitEvent As Long
 
  pid = Shell(App.Path & "\AlcorMP_698x2\loaddrv.exe uninstall_058F6387")
  hProcess = OpenProcess(SYNCHRONIZE + PROCESS_QUERY_INFORMATION + PROCESS_TERMINATE, 0, pid)
  ExitEvent = WaitForSingleObject(hProcess, INFINITE)
  Call CloseHandle(hProcess)
  KillProcess ("LoadDrv.exe")
 
 
                        
 '=========================================
 '    POWER on
 '=========================================
'T1:
RW_Test_Label:
 
 If MPFlag = 1 Then
         Call PowerSet(3)
          cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        
         Call MsecDelay(0.5)  ' power of to unload MPDriver

           cardresult = DO_WritePort(card, Channel_P1A, &HFB)  'sel socket
           Call PowerSet(1500)
     
        
         Call MsecDelay(1.2)
        MPFlag = 0
 Else
          cardresult = DO_WritePort(card, Channel_P1A, &HFB)
         Call PowerSet(1500)
         
         Call MsecDelay(1.2)
End If
         Call LoadRWTest_Click_AU6992

        
        
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
     '   GoTo T2
       If PassTime > 5 Then
           TestResult = "Bin3"
           MPTester.TestResultLab = "Bin3:RW Ready Fail"
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
       
            Exit Sub
       End If
         
T2:
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        Call StartRWTest_Click_AU6988
        
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
                Or AlcorMPMessage = WM_FT_CHECK_CERBGPO_FAIL _
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
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
            
              cardresult = DO_WritePort(card, Channel_P1A, &HFB)  ' power off
            
             Call PowerSet(1500)
        
            
       
            Exit Sub
        End If
        
        
     
               
        Select Case AlcorMPMessage
        
        Case WM_FT_RW_UNKNOW_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:UnKnow Fail"
          
             ContFail = ContFail + 1
        
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
             
        Case WM_FT_RW_ROM_FAIL
              TestResult = "Bin4"
              MPTester.TestResultLab = "Bin4:ROM FAIL "
              ContFail = ContFail + 1
              
        Case WM_FT_RW_RAM_FAIL
              TestResult = "Bin5"
              MPTester.TestResultLab = "Bin5:RAM FAIL "
               ContFail = ContFail + 1
        Case WM_FT_RW_RW_PASS
        
               
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
              
               End If
               
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:Undefine Fail"
          
             ContFail = ContFail + 1
        
               
        End Select
                               
       cardresult = DO_WritePort(card, Channel_P1A, &HFB)
         Call PowerSet(1500)

End Sub

Public Sub AU6992EQCHWSingleTestSub()
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
 'Dim ReMP_Flag As Byte
 Dim HV_Result As String
 Dim LV_Result As String
 
 
MPTester.TestResultLab = ""
HV_Result = ""
LV_Result = ""
EQC_HV = False
EQC_LV = False

 
   MPTester.TestResultLab = ""
'===============================================================
' Fail location initial
'===============================================================
 
If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
    FileCopy App.Path & "\AlcorMP_698x2\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
    Call MsecDelay(5)
End If


NewChipFlag = 0
If OldChipName <> ChipName Then
            FileCopy App.Path & "\AlcorMP_698x2\New_INI\" & ChipName & "\ROM\ROM.Hex", App.Path & "\AlcorMP_698x2\ROM.Hex"
            FileCopy App.Path & "\AlcorMP_698x2\New_INI\" & ChipName & "\RAM\\RAM.Bin", App.Path & "\AlcorMP_698x2\RAM.Bin"
            FileCopy App.Path & "\AlcorMP_698x2\New_INI\" & ChipName & "\AlcorMP.ini", App.Path & "\AlcorMP_698x2\AlcorMP.ini"
            FileCopy App.Path & "\AlcorMP_698x2\New_INI\" & ChipName & "\PE.bin", App.Path & "\AlcorMP_698x2\PE.bin"
            FileCopy App.Path & "\AlcorMP_698x2\New_INI\" & ChipName & "\FT.ini", App.Path & "\FT.ini"
            FileCopy App.Path & "\AlcorMP_698x2\New_INI\" & ChipName & "\FT.ini", App.Path & "\AlcorMP_698x2\FT.ini"
            NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName
 

'==============================================================
' when begin RW Test, must clear MP program
'===============================================================


'(1)  /// for Auto mode
winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
  Loop While winHwnd <> 0
End If

'(1)
winHwnd = FindWindow(vbNullString, AU6992MPCaption)
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, AU6992MPCaption)
  Loop While winHwnd <> 0
End If
'(2)
winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
  Loop While winHwnd <> 0
End If
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail
 
 
 '====================================
 '  Fix Card
 '====================================
' GoTo T1
 If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag = True) Then
 
 
   If MPTester.NoMP.Value = 1 Then
        
        If (NewChipFlag = 0) And (MPTester.Check1.Value = 0) Then  ' force condition
           GoTo RW_Test_Label
        End If
    End If
       
    
    If MPTester.ResetMPFailCounter.Value = 1 Then
       ContFail = 0
    End If
    
 '==============================================================
' when begin MP, must close RW program
'===============================================================
   MPFlag = 1
 
    winHwnd = FindWindow(vbNullString, "UFD Test")
    If winHwnd <> 0 Then
      Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(1.2)
        winHwnd = FindWindow(vbNullString, "UFD Test")
      Loop While winHwnd <> 0
    End If
 
       '  power on
       cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        Call PowerSet(3)   ' close power to disable chip
       Call MsecDelay(0.5)  ' power for load MPDriver
       MPTester.Print "wait for MP Ready"
       Call LoadMP_Click_AU6992
 
        OldTimer = Timer
        AlcorMPMessage = 0
        Debug.Print "begin"
        Do
           ' DoEvents
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
            PassTime = Timer - OldTimer
            'Debug.Print AlcorMPMessage
        Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 30 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3
        '====================================================
        If PassTime > 30 Then
          '(1)
           MPTester.Print "MP Ready Fail"
            winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
              Loop While winHwnd <> 0
            End If
           '(2)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
            If winHwnd <> 0 Then
              Do
              rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
              Call MsecDelay(0.5)
              winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
              Loop While winHwnd <> 0
            End If
            
        
             MPTester.TestResultLab = "Bin3:MP Ready Fail"
             TestResult = "Bin3"
             MPTester.Print "MP Ready Fail"
     
              
            Exit Sub
        End If
        
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_MP_START Then
         
             
            cardresult = DO_WritePort(card, Channel_P1A, &HFB)  ' sel chip
            'Call PowerSet(500)   ' close power to disable chip
            Call PowerSet2(1, "3.3", "0.5", 1, "3.3", "0.5", 1)
 
          
             
            Dim TimerCounter As Integer
            Dim TmpString As String
            
             
            Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName("vid")
            Loop While TmpString = "" And TimerCounter < 150
             
            Call MsecDelay(0.3)
            
             If TmpString = "" Then   ' can not find device after 15 s
             
               TestResult = "Bin2"
               MPTester.TestResultLab = "Bin2:MP UNKNOW Fail when enter MP"
               Exit Sub
             End If
             
             Call MsecDelay(2.5)
               
             MPTester.Print " MP Begin....."
             
             Call StartMP_Click_AU6992
   
             
              OldTimer = Timer
              AlcorMPMessage = 0
              ReMP_Flag = 0
              
                Do
                   ' DoEvents
                    If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                        AlcorMPMessage = mMsg.message
                        TranslateMessage mMsg
                        DispatchMessage mMsg
                        
                        If (AlcorMPMessage = WM_FT_MP_FAIL) And (ReMP_Flag = 0) Then
                            'ReMP_Flag = 1
                            AlcorMPMessage = 1
                            cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'close power
                            Call MsecDelay(0.3)
                            cardresult = DO_WritePort(card, Channel_P1A, &HFB)  'open power
                            Call MsecDelay(2.2)
                            Call RefreshMP_Click_AU6992
                            Call MsecDelay(0.5)
                            Call StartMP_Click_AU6992
                            
                            ReMP_Counter = ReMP_Counter + 1
                            If ReMP_Counter >= ReMP_Limit Then
                                ReMP_Flag = 1
                                ReMP_Counter = 0
                            End If
                        End If
                    
                    End If
                    PassTime = Timer - OldTimer
                Loop Until AlcorMPMessage = WM_FT_MP_PASS _
                Or AlcorMPMessage = WM_FT_MP_FAIL _
                Or AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL _
                Or PassTime > 65 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                
                If AlcorMPMessage = WM_FT_MP_PASS Then
                    If ReMP_Flag = 0 Then
                        MsecDelay (MPIdleTime * (ReMP_Limit - ReMP_Counter))
                    End If
                    ReMP_Counter = 0
                End If
                
                MPTester.Print "MP work time="; PassTime
                 MPTester.MPText.Text = Hex(AlcorMPMessage)
                '================================================
                '  Handle MP work time out error
                '===============================================
                
               ' time out fail
                If PassTime > 65 Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    '(1)
                    winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                      Loop While winHwnd <> 0
                    End If
                    
                    '(2)
                      winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    Exit Sub
                End If
                
                ' MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    
                    winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                      Loop While winHwnd <> 0
                    End If
                    
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    
                    Exit Sub
                End If
                
                
                 'unknow fail
                If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                    MPTester.Print "MP UNKNOW Fail"
                    
                   winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                    If winHwnd <> 0 Then
                     Do
                     rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                     Call MsecDelay(0.5)
                     winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                     Loop While winHwnd <> 0
                     End If
                     
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                     If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                     End If
                     
                     
                     
                    Exit Sub
                End If
                 
                
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_PASS Then
                     MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPTester.Print "MP PASS"
                End If
        End If
   
End If
'=========================================
 '    Close MP program
 '=========================================
winHwnd = FindWindow(vbNullString, AU6992MPCaption)
If winHwnd <> 0 Then
  Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, AU6992MPCaption)
  Loop While winHwnd <> 0
    Call MsecDelay(0.2)
    KillProcess ("AlcorMP.exe")

End If
    
 winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
    Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
    Loop While winHwnd <> 0
End If

  Dim pid As Long          ' unload driver
  Dim hProcess As Long
  Dim ExitEvent As Long
 
  pid = Shell(App.Path & "\AlcorMP_698x2\loaddrv.exe uninstall_058F6387")
  hProcess = OpenProcess(SYNCHRONIZE + PROCESS_QUERY_INFORMATION + PROCESS_TERMINATE, 0, pid)
  ExitEvent = WaitForSingleObject(hProcess, INFINITE)
  Call CloseHandle(hProcess)
  KillProcess ("LoadDrv.exe")
 
 
                        
 '=========================================
 '    POWER on
 '=========================================
'T1:
RW_Test_Label:
 
If MPFlag = 1 Then
        
    If (EQC_HV = False) And (EQC_LV = False) Then
        Call PowerSet2(1, "3.6", "0.15", 1, "3.6", "0.15", 1)
        cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        Call MsecDelay(0.5)  ' power of to unload MPDriver

        cardresult = DO_WritePort(card, Channel_P1A, &HFB)  'sel socket
        Call MsecDelay(1.3)
        'WaitDevOn ("vid_058f")
        'Call MsecDelay(0.2)
        EQC_HV = True
    End If
     
    MPFlag = 0
 
Else
          
    If (EQC_HV = False) And (EQC_LV = False) Then
        Call PowerSet2(1, "3.6", "0.15", 1, "3.6", "0.15", 1)
        cardresult = DO_WritePort(card, Channel_P1A, &HFB)
        Call MsecDelay(1.3)
        'WaitDevOn ("pid_058f")
        'Call MsecDelay(0.2)
        EQC_HV = True
    End If

End If
         Call LoadRWTest_Click_AU6992
        
                
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
     '   GoTo T2
       If PassTime > 5 Then
           TestResult = "Bin3"
           MPTester.TestResultLab = "Bin3:RW Ready Fail"
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
       
            Exit Sub
       End If
         
T2:
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        Call StartRWTest_Click_AU6988
        
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
            Or AlcorMPMessage = WM_FT_CHECK_CERBGPO_FAIL _
            Or AlcorMPMessage = WM_FT_CHECK_HW_CODE_FAIL _
            Or AlcorMPMessage = WM_FT_PHYREAD_FAIL _
            Or AlcorMPMessage = WM_FT_ECC_FAIL _
            Or AlcorMPMessage = WM_FT_NOFREEBLOCK_FAIL _
            Or AlcorMPMessage = WM_FT_LODECODE_FAIL _
            Or AlcorMPMessage = WM_FT_RELOADCODE_FAIL _
            Or AlcorMPMessage = WM_FT_TESTUNITREADY_FAIL _
            Or PassTime > 10 _
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
            MPTester.TestResultLab = "HV: UnKnow Fail"
            'ContFail = ContFail + 1
        
        Case WM_FT_CHECK_HW_CODE_FAIL
             TestResult = "Bin5"
             MPTester.TestResultLab = "HV: HW-ID Fail"
             'ContFail = ContFail + 1
        
        Case WM_FT_TESTUNITREADY_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "HV: TestUnitReady Fail"
             'ContFail = ContFail + 1
        
        Case WM_FT_RW_SPEED_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "HV: SPEED Error "
             'ContFail = ContFail + 1

        Case WM_FT_RW_RW_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = "HV: RW FAIL "
             'ContFail = ContFail + 1

        Case WM_FT_CHECK_CERBGPO_FAIL

            TestResult = "Bin3"
            MPTester.TestResultLab = "HV: GPO/RB FAIL "
            'ContFail = ContFail + 1

        Case WM_FT_RW_ROM_FAIL
            TestResult = "Bin4"
            MPTester.TestResultLab = "HV: ROM FAIL "
            'ContFail = ContFail + 1

        Case WM_FT_RW_RAM_FAIL, WM_FT_PHYREAD_FAIL, WM_FT_ECC_FAIL, WM_FT_NOFREEBLOCK_FAIL, WM_FT_LODECODE_FAIL, WM_FT_RELOADCODE_FAIL
            TestResult = "Bin5"
            MPTester.TestResultLab = "HV: RAM FAIL "
            'ContFail = ContFail + 1

        Case WM_FT_RW_RW_PASS

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
                'ContFail = 0 '
            Else

                TestResult = "Bin3"

            End If
                       
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = "HV: Undefine Fail"

             ContFail = ContFail + 1
        
               
        End Select
        
        HV_Result = TestResult
        TestResult = ""
        
        cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'Power OFF UNLoad Device
        Call MsecDelay(0.2)
        Call PowerSet2(1, "3.0", "0.15", 1, "3.0", "0.15", 1)
        EQC_LV = True
        cardresult = DO_WritePort(card, Channel_P1A, &HFB)  'Power ON
        Call MsecDelay(1.3)
        'WaitDevOn ("vid_058f")
        'Call MsecDelay(0.2)
        GoTo RW_Test_Label
        
    ElseIf (EQC_HV = True) And (EQC_LV = True) Then
        
        Select Case AlcorMPMessage
  
        Case WM_FT_RW_UNKNOW_FAIL
            TestResult = "Bin2"
            MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: UnKnow Fail"
            'ContFail = ContFail + 1
        
        Case WM_FT_CHECK_HW_CODE_FAIL
             TestResult = "Bin5"
             MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: HW-ID Fail"
             'ContFail = ContFail + 1
        
        Case WM_FT_TESTUNITREADY_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: TestUnitReady Fail"
             'ContFail = ContFail + 1
        
        Case WM_FT_RW_SPEED_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: SPEED Error "
             'ContFail = ContFail + 1

        Case WM_FT_RW_RW_FAIL
             TestResult = "Bin3"
             MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: RW FAIL "
             'ContFail = ContFail + 1

        Case WM_FT_CHECK_CERBGPO_FAIL

            TestResult = "Bin3"
            MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: GPO/RB FAIL "
            'ContFail = ContFail + 1

        Case WM_FT_RW_ROM_FAIL
            TestResult = "Bin4"
            MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: ROM FAIL "
            'ContFail = ContFail + 1

        Case WM_FT_RW_RAM_FAIL, WM_FT_PHYREAD_FAIL, WM_FT_ECC_FAIL, WM_FT_NOFREEBLOCK_FAIL, WM_FT_LODECODE_FAIL, WM_FT_RELOADCODE_FAIL
            TestResult = "Bin5"
            MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: RAM FAIL "
            'ContFail = ContFail + 1

        Case WM_FT_RW_RW_PASS

            For LedCount = 1 To 20
                Call MsecDelay(0.1)
                cardresult = DO_ReadPort(card, Channel_P1B, LightOn)
                
                If LightOn = &HEF Or LightOn = &HCF Or LightOn = 223 Then
                    Exit For
                End If
            Next LedCount

            MPTester.Print "light="; LightOn

            If (LightOn = &HEF Or LightOn = &HCF Or LightOn = 223) Then
                MPTester.TestResultLab = MPTester.TestResultLab & vbCrLf & "LV: PASS"
                TestResult = "PASS"
                'ContFail = 0 '
            Else

                TestResult = "Bin3"

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
            TestResult = "PASS"
            ContFail = 0
        Else
            TestResult = "Bin2"
            ContFail = ContFail + 1
        End If
        
        cardresult = DO_WritePort(card, Channel_P1A, &HFB)
    End If

End Sub

Public Sub LoadMP_Click_AU6992()

Dim TimePass
Dim rt2
    ' find window1
'    winHwnd = FindWindow(vbNullString, "Module Update")
 
    ' run program
'    If winHwnd = 0 Then
'        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\AlcorMP_698x2\AU699x_MP_Update_Patch_v09.07.06.01-FT.exe", "", "", SW_SHOW)
'    End If

'    Do
'        winHwnd = FindWindow(vbNullString, "Module Update")
'    Loop While winHwnd = 0
'
'    winHwnd = FindWindow(vbNullString, "Module Update")
'    Do
'        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
'        winHwnd = FindWindow(vbNullString, "Module Update")
'    Loop While winHwnd <> 0
 
    winHwnd = FindWindow(vbNullString, AU6992MPCaption1)        ' AU6992MPCaption1 = 698x UFD MP
 
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\AlcorMP_698x2\AlcorMP.exe", "", "", SW_SHOW)
    End If

    SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags
    
End Sub

Public Sub LoadNewMP_Click_AU6996Flash()

Dim TimePass
Dim rt2
    ' find window1
'    winHwnd = FindWindow(vbNullString, "Module Update")
 
    ' run program
'    If winHwnd = 0 Then
'        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\AlcorMP_6996Flash\AU6996_FT_MP_TOOL_v10.03.10.01-6996.exe", "", "", SW_SHOW)
'    End If

'    Do
'        winHwnd = FindWindow(vbNullString, "Module Update")
'    Loop While winHwnd = 0
'
'    winHwnd = FindWindow(vbNullString, "Module Update")
'    Do
'        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
'        winHwnd = FindWindow(vbNullString, "Module Update")
'    Loop While winHwnd <> 0
 
    winHwnd = FindWindow(vbNullString, AU6996MPCaption1)
 
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\AlcorMP_6996Flash\AlcorMP.exe", "", "", SW_SHOW)
    End If
    
    SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags

End Sub
 
Public Sub LoadMP_Click_AU6996Reader()

Dim TimePass
Dim rt2
    ' find window1
'    winHwnd = FindWindow(vbNullString, "Module Update")
'
'    ' run program
'    If winHwnd = 0 Then
'        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\AlcorMP_6996Reader\AU6996_FT_MP_TOOL_v10.03.16.01-6996.exe", "", "", SW_SHOW)
'    End If
'
'    Do
'        winHwnd = FindWindow(vbNullString, "Module Update")
'    Loop While winHwnd = 0
'
'    winHwnd = FindWindow(vbNullString, "Module Update")
'    Do
'        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
'        winHwnd = FindWindow(vbNullString, "Module Update")
'    Loop While winHwnd <> 0
 
    winHwnd = FindWindow(vbNullString, AU6996MPCaption1)
 
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\AlcorMP_6996Reader\AlcorMP.exe", "", "", SW_SHOW)
    End If

    SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags
    
End Sub
 
Public Sub LoadMP_Click_AU6997()

Dim TimePass
Dim rt2
    winHwnd = FindWindow(vbNullString, AU6997MPCaption1)
 
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\AU6997FT\AlcorMP.exe", "", "", SW_SHOW)
    End If
    
    SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags

End Sub
 
Public Sub LoadMP_Click_AU6988_20090904()

Dim TimePass
Dim rt2
    ' find window
'    winHwnd = FindWindow(vbNullString, "Module Update")
'
'    ' run program
'    If winHwnd = 0 Then
'        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\AU6986AU6988_20090904\AU6986,6988_FT_Test_MP_Patch_v09.09.04.01-FT.exe", "", "", SW_SHOW)
'    End If
'
'    Do
'        winHwnd = FindWindow(vbNullString, "Module Update")
'    Loop While winHwnd = 0
'
'    winHwnd = FindWindow(vbNullString, "Module Update")
'    Do
'        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
'        winHwnd = FindWindow(vbNullString, "Module Update")
'    Loop While winHwnd <> 0
 
    winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
 
    ' run program
    If winHwnd = 0 Then
    Call ShellExecute(MPTester.hwnd, "open", App.Path & "\AU6986AU6988_20090904\AlcorMP.exe", "", "", SW_SHOW)
    End If

    SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags
    
End Sub
 
Public Sub LoadMP_Click_AU6988()

Dim TimePass
Dim rt2
    ' find window
'    winHwnd = FindWindow(vbNullString, "Module Update")
'
'    ' run program
'    If winHwnd = 0 Then
'        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\AlcorMP_698x_PD\AU698x_FT_Test_MP_Tool_v09.01.06.01-FT.exe", "", "", SW_SHOW)
'    End If
'
'    Do
'        winHwnd = FindWindow(vbNullString, "Module Update")
'    Loop While winHwnd = 0
'
'    winHwnd = FindWindow(vbNullString, "Module Update")
'    Do
'        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
'        winHwnd = FindWindow(vbNullString, "Module Update")
'    Loop While winHwnd <> 0

    winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
 
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\AlcorMP_698x_PD\AlcorMP.exe", "", "", SW_SHOW)
    End If

    SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags
    
End Sub
Public Sub LoadMP_Click_AU6988D53()

Dim TimePass
Dim rt2
    ' find window
'    winHwnd = FindWindow(vbNullString, "Module Update")
'
'    ' run program
'    If winHwnd = 0 Then
'        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\AlcorMP_698x_PD\AU698x_FT_Test_MP_Tool_v09.01.06.01-FT.exe", "", "", SW_SHOW)
'    End If
'
'    Do
'        winHwnd = FindWindow(vbNullString, "Module Update")
'    Loop While winHwnd = 0
'
'    winHwnd = FindWindow(vbNullString, "Module Update")
'    Do
'        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
'        winHwnd = FindWindow(vbNullString, "Module Update")
'    Loop While winHwnd <> 0

    winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
 
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\AlcorMP_698x_PD\AlcorMPNew.exe", "", "", SW_SHOW)
    End If

    SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags
    
End Sub

Public Sub LoadRWTest_Click_AU6988_20090904()

Dim TimePass
    ' find window
    winHwnd = FindWindow(vbNullString, "UFD Test")
     
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\AU6986AU6988_20090904\UFDTest.exe", "", "", SW_SHOW)
    End If
 
End Sub
Public Sub LoadRWTest_Click_AU6988_K9F1G()

Dim TimePass
    ' find window
    winHwnd = FindWindow(vbNullString, "UFD Test")
 
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\AlcorMP_6988\UFDTest.exe", "6988", "", SW_SHOW)
    End If
 
End Sub

Public Sub LoadRWTest_Click_AU6988()

Dim TimePass
    ' find window
    winHwnd = FindWindow(vbNullString, "UFD Test")
     
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\AlcorMP_698x_PD\UFDTest.exe", "", "", SW_SHOW)
    End If
 
End Sub

Public Sub LoadRWTest_Click_AU6992()

Dim TimePass
    ' find window
    winHwnd = FindWindow(vbNullString, "UFD Test")
     
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\AlcorMP_698x2\UFDTest.exe", "6992", "", SW_SHOW)
    End If
 
End Sub
 
Public Sub LoadRWTest_Click_AU6996Flash()

Dim TimePass
    ' find window
    winHwnd = FindWindow(vbNullString, "UFD Test")
     
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\AlcorMP_6996Flash\UFDTest.exe", "", "", SW_SHOW)
    End If
 
End Sub
 
Public Sub LoadRWTest_Click_AU6997()

Dim TimePass
    ' find window
    winHwnd = FindWindow(vbNullString, "UFD Test")
     
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\AU6997FT\UFDTest.exe", "", "", SW_SHOW)
    End If
 
End Sub
 
Public Sub LoadRWTest_Click_AU6996Reader()

Dim TimePass
    ' find window
    winHwnd = FindWindow(vbNullString, "UFD Test")
     
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\AlcorMP_6996Reader\UFDTest.exe", "", "", SW_SHOW)
    End If
 
End Sub

Public Sub LoadRWTest_Click_AU6990()

Dim TimePass
    ' find window
    winHwnd = FindWindow(vbNullString, "UFD Test")
 
    ' run program
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & "\AlcorMP_698x2\UFDTest.exe", "6990", "", SW_SHOW)
    End If
 
End Sub

Public Sub StartRWTest_Click_AU6988()

Dim rt2
    winHwnd = FindWindow(vbNullString, "UFD Test")
    Debug.Print "WindHandle="; winHwnd
    rt2 = PostMessage(winHwnd, WM_FT_RW_START, 0&, 0&)

End Sub
 
Public Sub StartMP_Click_AU6988()

Dim rt2
    winHwnd = FindWindow(vbNullString, "Alcor Micro UFD Manufacture Program")
    Debug.Print "WindHandle="; winHwnd
    rt2 = PostMessage(winHwnd, WM_FT_MP_START, 0&, 0&)

End Sub

Public Sub StartMP_Click_AU6992()

Dim rt2
    winHwnd = FindWindow(vbNullString, AU6992MPCaption)
    Debug.Print "WindHandle="; winHwnd
    rt2 = PostMessage(winHwnd, WM_FT_MP_START, 0&, 0&)
    
End Sub

Public Sub RefreshMP_Click_AU6992()

Dim rt2
    winHwnd = FindWindow(vbNullString, AU6992MPCaption)
    Debug.Print "WindHandle="; winHwnd
    rt2 = PostMessage(winHwnd, WM_FT_MP_REFRESH, 0&, 0&)
    
End Sub

Public Sub RefreshMP_Click_AU6996()

Dim rt2
    winHwnd = FindWindow(vbNullString, AU6996MPCaption)
    Debug.Print "WindHandle="; winHwnd
    rt2 = PostMessage(winHwnd, WM_FT_MP_REFRESH, 0&, 0&)
    
End Sub

Public Sub StartMP_Click_AU6996()

Dim rt2
    Do
        Call MsecDelay(2#)
        winHwnd = FindWindow(vbNullString, AU6996MPCaption)
        Debug.Print "MP WindHandle="; winHwnd
    Loop Until winHwnd <> 0
    
    rt2 = PostMessage(winHwnd, WM_FT_MP_START, 0&, 0&)
    Debug.Print "Send MP start message"

End Sub
 
Public Sub StartMP_Click_AU6997()
 
Dim rt2
    Do
        Call MsecDelay(2#)
        winHwnd = FindWindow(vbNullString, AU6997MPCaption)
        Debug.Print "MP WindHandle="; winHwnd
    Loop Until winHwnd <> 0
    
    rt2 = PostMessage(winHwnd, WM_FT_MP_START, 0&, 0&)
    Debug.Print "Send MP start message"
    
End Sub

Public Sub AutoMP_sub()

Dim tmpName As String
Dim CurrentICMPCaption As String
Dim CurrentICMPCaption1 As String
Dim mMsg As MSG
Dim rt2
Dim PassTime
Dim OldTimer
Dim TimerCounter As Integer
Dim TmpString As String
    
    ReMP_Flag = 0
    ReMP_Counter = 0
    If OldChipName = "" Then
        MsgBox ("Please Send Test Signal Form HOST First!")
        Exit Sub
    End If
    
    If InStr(OldChipName, "AU6992") Then
        CurDevicePar.ShortName = "AU6992"
        CurDevicePar.DeviceFolder = "\AlcorMP_698x2"
        CurDevicePar.MP_LoadTitle = "698x UFD MP"
        CurDevicePar.MP_WorkTitle = "698x UFD MP, Cycle Time : 33 ns"
        CurDevicePar.MP_ToolFileName = "AlcorMP.exe"
        CurDevicePar.FT_ToolFileName = "UFDTest.exe"
        CurDevicePar.FT_ToolTitle = "UFD Test"
        CurDevicePar.Exec_Par1 = "6992"
        CurDevicePar.FullName = "AU6992R53HLF23"
        CurDevicePar.SetStdV1 = "5.0"
        CurDevicePar.SetStdV2 = "5.0"
        CurDevicePar.SetStdI1 = "0.2"
        CurDevicePar.SetStdI2 = "0.2"
    ElseIf InStr(OldChipName, "AU6988") Then
        CurDevicePar.ShortName = "AU6988"
        CurDevicePar.DeviceFolder = "\AlcorMP_6988"
        CurDevicePar.MP_LoadTitle = "698x UFD MP"
        CurDevicePar.MP_WorkTitle = "698x UFD MP, Cycle Time : 33 ns"
        CurDevicePar.MP_ToolFileName = "AlcorMP.exe"
        CurDevicePar.FT_ToolFileName = "UFDTest.exe"
        CurDevicePar.FT_ToolTitle = "UFD Test"
        CurDevicePar.Exec_Par1 = "6988"
        CurDevicePar.FullName = "AU6988D52HLF2I"
        CurDevicePar.SetStdV1 = "5.0"
        CurDevicePar.SetStdV2 = "5.0"
        CurDevicePar.SetStdI1 = "0.2"
        CurDevicePar.SetStdI2 = "0.2"
    ElseIf InStr(OldChipName, "AU6996") Then
        CurDevicePar.ShortName = "AU6996"
        CurDevicePar.DeviceFolder = "\AlcorMP_6996Flash"
        CurDevicePar.MP_LoadTitle = "698x UFD MP"
        CurDevicePar.MP_WorkTitle = "698x UFD MP, Cycle Time : 33 ns"
        CurDevicePar.MP_ToolFileName = "AlcorMP.exe"
        CurDevicePar.FT_ToolFileName = "UFDTest.exe"
        CurDevicePar.FT_ToolTitle = "UFD Test"
        CurDevicePar.Exec_Par1 = "6996"
        CurDevicePar.FullName = "AU6996A51ILF2B"
        CurDevicePar.SetStdV1 = "5.0"
        CurDevicePar.SetStdV2 = "5.0"
        CurDevicePar.SetStdI1 = "0.2"
        CurDevicePar.SetStdI2 = "0.2"
    End If
    
    '(1)  /// for Auto mode
    winHwnd = FindWindow(vbNullString, CurDevicePar.MP_LoadTitle)
    If winHwnd <> 0 Then
        Do
            rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
            Call MsecDelay(0.5)
            winHwnd = FindWindow(vbNullString, CurDevicePar.MP_LoadTitle)
        Loop While winHwnd <> 0
    End If
    
    '(1)
    winHwnd = FindWindow(vbNullString, CurDevicePar.MP_WorkTitle)
    If winHwnd <> 0 Then
        Do
            rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
            Call MsecDelay(0.5)
            winHwnd = FindWindow(vbNullString, CurDevicePar.MP_WorkTitle)
        Loop While winHwnd <> 0
    End If
    
    '(2)
    winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
    If winHwnd <> 0 Then
        Do
            rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
            Call MsecDelay(0.5)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
        Loop While winHwnd <> 0
    End If
     
    '==============================================================
    ' when begin MP, must close RW program
    '===============================================================
    MPFlag = 1
     
    winHwnd = FindWindow(vbNullString, "UFD Test")
    If winHwnd <> 0 Then
        Do
            rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
            Call MsecDelay(0.5)
            winHwnd = FindWindow(vbNullString, "UFD Test")
        Loop While winHwnd <> 0
    End If
     
    '  power off
    cardresult = DO_WritePort(card, Channel_P1A, &HFF)
    Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
    Call MsecDelay(0.5)  ' power for load MPDriver
    MPTester.Print "wait for MP Ready"
    
'    winHwnd = FindWindow(vbNullString, "Module Update")
'    ' run program
'    If winHwnd = 0 Then
'        Call ShellExecute(MPTester.hwnd, "open", App.Path & CurDevicePar.DeviceFolder & "\" & CurDevicePar.UpdateModuleName, "", "", SW_SHOW)
'    End If
'
'    Do
'        winHwnd = FindWindow(vbNullString, "Module Update")
'    Loop While winHwnd = 0

AutoMP_NoUpdate:

'    winHwnd = FindWindow(vbNullString, "Module Update")
'    Do
'        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
'        winHwnd = FindWindow(vbNullString, "Module Update")
'    Loop While winHwnd <> 0
     
    winHwnd = FindWindow(vbNullString, CurDevicePar.MP_LoadTitle)
    If winHwnd = 0 Then
        Call ShellExecute(MPTester.hwnd, "open", App.Path & CurDevicePar.DeviceFolder & "\" & CurDevicePar.MP_ToolFileName, "", "", SW_SHOW)
    End If
    
    SetWindowPos winHwnd, HWND_TOPMOST, 300, 300, 0, 0, Flags
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
    Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 30 _
    Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                  
    MPTester.Print "Ready Time="; PassTime
            
    '====================================================
    '  handle MP load time out
    '====================================================
    
    If PassTime > 30 Then
    '(1)
        MPTester.Print "MP Ready Fail"
        winHwnd = FindWindow(vbNullString, CurDevicePar.MP_LoadTitle)
        If winHwnd <> 0 Then
            Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, CurDevicePar.MP_LoadTitle)
            Loop While winHwnd <> 0
        End If
        MPTester.TestResultLab = "MP Ready Fail"
        MPTester.Print "MP Ready Fail"
        Exit Sub
    End If
            
    '====================================================
    '  MP begin
    '====================================================
            
    If AlcorMPMessage = WM_FT_MP_START Then
             
        If Dual_Flag = True Then
            cardresult = DO_WritePort(card, Channel_P1A, &HFD)  ' sel chip
        Else
            cardresult = DO_WritePort(card, Channel_P1A, &HFB)  ' sel chip
        End If
        
        If Dir("D:\LABPC.PC") = "LABPC.PC" Then
            Call PowerSet2(1, "5.0", "0.5", 1, "5.0", "0.5", 1)
        Else
            If EQC_Flag Then
                Call PowerSet2(1, CurDevicePar.SetHLVStd1, "0.5", 1, CurDevicePar.SetHLVStd2, "0.5", 1)
            Else
                Call PowerSet2(1, CurDevicePar.SetStdV1, "0.5", 1, CurDevicePar.SetStdV2, "0.5", 1)
            End If
        End If
                        
        If MPTester.AutoMP_Option.Value <> 1 Then
            MPTester.Print "Load MP Success"
            Exit Sub
        End If
                
        Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName("vid")
        Loop While TmpString = "" And TimerCounter < 150
                          
        Call MsecDelay(0.3)
                 
        If TmpString = "" Then   ' can not find device after 15 s
            MPTester.TestResultLab = "Bin2:MP UNKNOW Fail when enter MP"
            Exit Sub
        End If
                 
        Call MsecDelay(2.5)
                   
        MPTester.Print " MP Begin....."
        ' ============= Push MP_Start Button =============
        winHwnd = FindWindow(vbNullString, CurDevicePar.MP_WorkTitle)
        rt2 = PostMessage(winHwnd, WM_FT_MP_START, 0&, 0&)
        '=================================================
                                
        ReMP_Flag = 0
        OldTimer = Timer
        AlcorMPMessage = 0
                    
        Do
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
                                
                If (AlcorMPMessage = WM_FT_MP_FAIL) And (ReMP_Flag = 0) Then
                    AlcorMPMessage = 1
                    cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'close power
                    Call MsecDelay(0.3)
                                
                    If Dual_Flag = True Then
                        cardresult = DO_WritePort(card, Channel_P1A, &HFD)  ' sel chip
                    Else
                        cardresult = DO_WritePort(card, Channel_P1A, &HFB)  ' sel chip
                    End If
                                
                    Call MsecDelay(2.2)
                    ' ============ Push MP_Refresh Button ============
                    winHwnd = FindWindow(vbNullString, CurDevicePar.MP_WorkTitle)
                    rt2 = PostMessage(winHwnd, WM_FT_MP_REFRESH, 0&, 0&)
                    ' ================================================
                    Call MsecDelay(0.5)
                    
                    ' ============= Push MP_Start Button =============
                    winHwnd = FindWindow(vbNullString, CurDevicePar.MP_WorkTitle)
                    rt2 = PostMessage(winHwnd, WM_FT_MP_START, 0&, 0&)
                    '=================================================
                                
                    ReMP_Counter = ReMP_Counter + 1
                    If ReMP_Counter >= ReMP_Limit Then
                        ReMP_Flag = 1
                        ReMP_Counter = 0
                    End If
                               
                End If
            End If
                        
            PassTime = Timer - OldTimer
                    
        Loop Until AlcorMPMessage = WM_FT_MP_PASS _
        Or AlcorMPMessage = WM_FT_MP_FAIL _
        Or AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL _
        Or PassTime > 65 _
        Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                           
        If AlcorMPMessage = WM_FT_MP_PASS Then
            ReMP_Counter = 0
        End If
                    
        MPTester.Print "MP work time="; PassTime
        MPTester.MPText.Text = Hex(AlcorMPMessage)
        
        '================================================
        '  Handle MP work time out error
        '===============================================
                    
        ' time out fail
        If PassTime > 65 Then
            MPTester.TestResultLab = "MP Time out Fail"
            MPTester.Print "MP Time out Fail"
            
            '(1)
            winHwnd = FindWindow(vbNullString, CurDevicePar.MP_WorkTitle)
            If winHwnd <> 0 Then
                Do
                    rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                    Call MsecDelay(0.5)
                    winHwnd = FindWindow(vbNullString, CurDevicePar.MP_WorkTitle)
                Loop While winHwnd <> 0
            End If
                        
            '(2)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
            If winHwnd <> 0 Then
                Do
                    rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                    Call MsecDelay(0.5)
                    winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                Loop While winHwnd <> 0
            End If
            
            Exit Sub
        End If
                    
        ' MP fail
        If AlcorMPMessage = WM_FT_MP_FAIL Then
            MPTester.TestResultLab = "Bin3:MP Function Fail"
            MPTester.Print "MP Function Fail"
            
            winHwnd = FindWindow(vbNullString, CurDevicePar.MP_WorkTitle)
            If winHwnd <> 0 Then
                Do
                    rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                    Call MsecDelay(0.5)
                    winHwnd = FindWindow(vbNullString, CurDevicePar.MP_WorkTitle)
                Loop While winHwnd <> 0
            End If
                        
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
            If winHwnd <> 0 Then
                Do
                    rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                    Call MsecDelay(0.5)
                    winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                Loop While winHwnd <> 0
            End If
                        
            Exit Sub
        End If
                             
        'unknow fail
        If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
            MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
            MPTester.Print "MP UNKNOW Fail"
            
            winHwnd = FindWindow(vbNullString, CurDevicePar.MP_WorkTitle)
            If winHwnd <> 0 Then
                Do
                    rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                    Call MsecDelay(0.5)
                    winHwnd = FindWindow(vbNullString, CurDevicePar.MP_WorkTitle)
                Loop While winHwnd <> 0
            End If
                         
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
            If winHwnd <> 0 Then
                Do
                    rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                    Call MsecDelay(0.5)
                    winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                Loop While winHwnd <> 0
            End If
            
            Exit Sub
        End If
                           
        ' mp pass
        If AlcorMPMessage = WM_FT_MP_PASS Then
            MPTester.TestResultLab = "MP PASS"
            MPTester.Print "MP PASS"
        End If
    End If
       
    '=========================================
     '    Close MP program
     '=========================================
    winHwnd = FindWindow(vbNullString, CurDevicePar.MP_WorkTitle)
    If winHwnd <> 0 Then
        Do
            rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
            Call MsecDelay(0.5)
            winHwnd = FindWindow(vbNullString, CurDevicePar.MP_WorkTitle)
        Loop While winHwnd <> 0
        
        Call MsecDelay(0.2)
        KillProcess ("AlcorMP.exe")
    End If
        
    winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
    If winHwnd <> 0 Then
        Do
            rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
            Call MsecDelay(0.5)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
        Loop While winHwnd <> 0
    End If

End Sub

Public Sub Close_FT_AP(Title As String)

Dim ProNum As Long

    ProNum = FindWindow(vbNullString, Title)
    If ProNum <> 0 Then
        Do
            ProNum = PostMessage(ProNum, WM_CLOSE, 0&, 0&)
            Call MsecDelay(0.2)
            ProNum = FindWindow(vbNullString, Title)
        Loop While ProNum <> 0
    End If
    
End Sub

Public Sub AU6992HWSingleTestSub()
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
 'Dim ReMP_Flag As Byte
   
   MPTester.TestResultLab = ""
'===============================================================
' Fail loacatio initial
'===============================================================
 
If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
    FileCopy App.Path & "\AlcorMP_698x2\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
    Call MsecDelay(5)
End If


NewChipFlag = 0
If OldChipName <> ChipName Then
            FileCopy App.Path & "\AlcorMP_698x2\New_INI\" & ChipName & "\ROM\ROM.Hex", App.Path & "\AlcorMP_698x2\ROM.Hex"
            FileCopy App.Path & "\AlcorMP_698x2\New_INI\" & ChipName & "\RAM\\RAM.Bin", App.Path & "\AlcorMP_698x2\RAM.Bin"
            FileCopy App.Path & "\AlcorMP_698x2\New_INI\" & ChipName & "\AlcorMP.ini", App.Path & "\AlcorMP_698x2\AlcorMP.ini"
            FileCopy App.Path & "\AlcorMP_698x2\New_INI\" & ChipName & "\PE.bin", App.Path & "\AlcorMP_698x2\PE.bin"
            FileCopy App.Path & "\AlcorMP_698x2\New_INI\" & ChipName & "\FT.ini", App.Path & "\FT.ini"
            FileCopy App.Path & "\AlcorMP_698x2\New_INI\" & ChipName & "\FT.ini", App.Path & "\AlcorMP_698x2\FT.ini"
            NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName
 

'==============================================================
' when begin RW Test, must clear MP porgram
'===============================================================


'(1)  /// for Auto mode
winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
  Loop While winHwnd <> 0
End If

'(1)
winHwnd = FindWindow(vbNullString, AU6992MPCaption)
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, AU6992MPCaption)
  Loop While winHwnd <> 0
End If
'(2)
winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
  Loop While winHwnd <> 0
End If
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail
 
 
 '====================================
 '  Fix Card
 '====================================
' GoTo T1
 If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag = True) Then
 
 
   If MPTester.NoMP.Value = 1 Then
        
        If (NewChipFlag = 0) And (MPTester.Check1.Value = 0) Then  ' force condition
           GoTo RW_Test_Label
        End If
    End If
       
    
    If MPTester.ResetMPFailCounter.Value = 1 Then
       ContFail = 0
    End If
    
 '==============================================================
' when begin MP, must close RW porgram
'===============================================================
   MPFlag = 1
 
    winHwnd = FindWindow(vbNullString, "UFD Test")
    If winHwnd <> 0 Then
      Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(1.2)
        winHwnd = FindWindow(vbNullString, "UFD Test")
      Loop While winHwnd <> 0
    End If
 
       '  power on
       cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        Call PowerSet(3)   ' close power to disable chip
       Call MsecDelay(0.5)  ' power for load MPDriver
       MPTester.Print "wait for MP Ready"
       Call LoadMP_Click_AU6992
 
        OldTimer = Timer
        AlcorMPMessage = 0
        Debug.Print "begin"
        Do
           ' DoEvents
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
            PassTime = Timer - OldTimer
            'Debug.Print AlcorMPMessage
        Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 30 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3
        '====================================================
        If PassTime > 30 Then
          '(1)
           MPTester.Print "MP Ready Fail"
            winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
              Loop While winHwnd <> 0
            End If
           '(2)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
            If winHwnd <> 0 Then
              Do
              rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
              Call MsecDelay(0.5)
              winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
              Loop While winHwnd <> 0
            End If
            
        
             MPTester.TestResultLab = "Bin3:MP Ready Fail"
             TestResult = "Bin3"
             MPTester.Print "MP Ready Fail"
     
              
            Exit Sub
        End If
        
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_MP_START Then
         
             
             cardresult = DO_WritePort(card, Channel_P1A, &HFB)  ' sel chip
              Call PowerSet(500)   ' close power to disable chip
             
 
          
             
            Dim TimerCounter As Integer
            Dim TmpString As String
            
             
            Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName("vid")
            Loop While TmpString = "" And TimerCounter < 150
             
            Call MsecDelay(0.3)
            
             If TmpString = "" Then   ' can not find device after 15 s
             
               TestResult = "Bin2"
               MPTester.TestResultLab = "Bin2:MP UNKNOW Fail when enter MP"
               Exit Sub
             End If
             
             Call MsecDelay(2.5)
               
             MPTester.Print " MP Begin....."
             
             Call StartMP_Click_AU6992
   
             
              OldTimer = Timer
              AlcorMPMessage = 0
              ReMP_Flag = 0
              
                Do
                   ' DoEvents
                    If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                        AlcorMPMessage = mMsg.message
                        TranslateMessage mMsg
                        DispatchMessage mMsg
                        
                        If (AlcorMPMessage = WM_FT_MP_FAIL) And (ReMP_Flag = 0) Then
                            'ReMP_Flag = 1
                            AlcorMPMessage = 1
                            cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'close power
                            Call MsecDelay(0.3)
                            cardresult = DO_WritePort(card, Channel_P1A, &HFB)  'open power
                            Call MsecDelay(2.2)
                            Call RefreshMP_Click_AU6992
                            Call MsecDelay(0.5)
                            Call StartMP_Click_AU6992
                            
                            ReMP_Counter = ReMP_Counter + 1
                            If ReMP_Counter >= ReMP_Limit Then
                                ReMP_Flag = 1
                                ReMP_Counter = 0
                            End If
                        End If
                    
                    End If
                    PassTime = Timer - OldTimer
                Loop Until AlcorMPMessage = WM_FT_MP_PASS _
                Or AlcorMPMessage = WM_FT_MP_FAIL _
                Or AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL _
                Or PassTime > 65 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                
                If AlcorMPMessage = WM_FT_MP_PASS Then
                    If ReMP_Flag = 0 Then
                        MsecDelay (MPIdleTime * (ReMP_Limit - ReMP_Counter))
                    End If
                    ReMP_Counter = 0
                End If
                
                MPTester.Print "MP work time="; PassTime
                 MPTester.MPText.Text = Hex(AlcorMPMessage)
                '================================================
                '  Handle MP work time out error
                '===============================================
                
               ' time out fail
                If PassTime > 65 Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    '(1)
                    winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                      Loop While winHwnd <> 0
                    End If
                    
                    '(2)
                      winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    Exit Sub
                End If
                
                ' MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    
                    winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                      Loop While winHwnd <> 0
                    End If
                    
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    
                    Exit Sub
                End If
                
                
                 'unknow fail
                 If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                     MPContFail = MPContFail + 1
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                    MPTester.Print "MP UNKNOW Fail"
                    
                   winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                    If winHwnd <> 0 Then
                     Do
                     rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                     Call MsecDelay(0.5)
                     winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                     Loop While winHwnd <> 0
                     End If
                     
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                     If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                     End If
                     
                     
                     
                    Exit Sub
                End If
                 
                
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_PASS Then
                     MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPTester.Print "MP PASS"
                End If
        End If
   
End If
'=========================================
 '    Close MP program
 '=========================================
winHwnd = FindWindow(vbNullString, AU6992MPCaption)
If winHwnd <> 0 Then
  Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, AU6992MPCaption)
  Loop While winHwnd <> 0
    Call MsecDelay(0.2)
    KillProcess ("AlcorMP.exe")

End If
    
 winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
    Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
    Loop While winHwnd <> 0
End If

  Dim pid As Long          ' unload driver
  Dim hProcess As Long
  Dim ExitEvent As Long
 
  pid = Shell(App.Path & "\AlcorMP_698x2\loaddrv.exe uninstall_058F6387")
  hProcess = OpenProcess(SYNCHRONIZE + PROCESS_QUERY_INFORMATION + PROCESS_TERMINATE, 0, pid)
  ExitEvent = WaitForSingleObject(hProcess, INFINITE)
  Call CloseHandle(hProcess)
  KillProcess ("LoadDrv.exe")
 
 
                        
 '=========================================
 '    POWER on
 '=========================================
'T1:
RW_Test_Label:
 
 If MPFlag = 1 Then
         Call PowerSet(3)
          cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        
         Call MsecDelay(0.5)  ' power of to unload MPDriver

           cardresult = DO_WritePort(card, Channel_P1A, &HFB)  'sel socket
           Call PowerSet(1500)
     
        
         Call MsecDelay(1.2)
        MPFlag = 0
 Else
          cardresult = DO_WritePort(card, Channel_P1A, &HFB)
         Call PowerSet(1500)
         
         Call MsecDelay(1.2)
End If
         Call LoadRWTest_Click_AU6992
        
                
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
     '   GoTo T2
       If PassTime > 5 Then
           TestResult = "Bin3"
           MPTester.TestResultLab = "Bin3:RW Ready Fail"
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
       
            Exit Sub
       End If
         
T2:
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        Call StartRWTest_Click_AU6988
        
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
            Or AlcorMPMessage = WM_FT_CHECK_CERBGPO_FAIL _
            Or AlcorMPMessage = WM_FT_CHECK_HW_CODE_FAIL _
            Or AlcorMPMessage = WM_FT_PHYREAD_FAIL _
            Or AlcorMPMessage = WM_FT_ECC_FAIL _
            Or AlcorMPMessage = WM_FT_NOFREEBLOCK_FAIL _
            Or AlcorMPMessage = WM_FT_LODECODE_FAIL _
            Or AlcorMPMessage = WM_FT_RELOADCODE_FAIL _
            Or AlcorMPMessage = WM_FT_TESTUNITREADY_FAIL _
            Or PassTime > 10 _
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
                cardresult = DO_WritePort(card, Channel_P1A, &HFA)  ' power off
                Exit Sub
            End If
        
        End If
     
               
        Select Case AlcorMPMessage
        
        Case WM_FT_RW_UNKNOW_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:UnKnow Fail"
             ContFail = ContFail + 1
        
        Case WM_FT_CHECK_HW_CODE_FAIL
             TestResult = "Bin5"
             MPTester.TestResultLab = "Bin5:HW-ID Fail"
             ContFail = ContFail + 1
        
        Case WM_FT_TESTUNITREADY_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:TestUnitReady Fail"
             ContFail = ContFail + 1

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
        
        Case WM_FT_RELOADCODE_FAIL
              TestResult = "Bin4"
              MPTester.TestResultLab = "Bin4:ReLoadCode FAIL "
              ContFail = ContFail + 1
        
        Case WM_FT_ECC_FAIL
              TestResult = "Bin5"
              MPTester.TestResultLab = "Bin5:ECC FAIL "
              ContFail = ContFail + 1
                    
        Case WM_FT_RW_RW_PASS
               
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
              
               End If
               
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:Undefine Fail"
          
             ContFail = ContFail + 1

        End Select
                               
       cardresult = DO_WritePort(card, Channel_P1A, &HFB)
         Call PowerSet(1500)

End Sub

Public Sub AU6992HWGOLDTestSub()
    
    If ChipName = "AU6992A51DLF2C" Then
        ChipName = "AU6992A51DLF23"
    ElseIf ChipName = "AU6992A52DLF2C" Then
        ChipName = "AU6992A52DLF23"
    ElseIf ChipName = "AU6992A53HLF2C" Then
        ChipName = "AU6992A53HLF23"
    ElseIf ChipName = "AU6992A54HLF2C" Then
        ChipName = "AU6992A54HLF23"
    ElseIf ChipName = "AU6992B54HLF2C" Then
        ChipName = "AU6992B54HLF23"
    ElseIf ChipName = "AU6992B53HLF2C" Then
        ChipName = "AU6992B53HLF23"
    ElseIf ChipName = "AU6992R53HLF2C" Then
        ChipName = "AU6992R53HLF23"
    ElseIf ChipName = "AU6992S53HLF2C" Then
        ChipName = "AU6992S53HLF23"
    Else
        ChipName = ChipName
    End If
    
    Call AU6992HW_DualTestSub
    
End Sub

Public Sub AU6992HW_DualTestSub()

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
 'Dim ReMP_Flag As Byte
 'Dim TmpChipName
 
 'TmpChipName = chipname
 'chipname = "AU6992A53DLF2A"
 
   MPTester.TestResultLab = ""
'===============================================================
' Fail loacatio initial
'===============================================================
 
If Dir("C:\WINDOWS\system32\drivers\mpfilt.sys") = "" Then
    FileCopy App.Path & "\AlcorMP_698x2\mpfilt.sys", "C:\WINDOWS\system32\drivers\mpfilt.sys"
    Call MsecDelay(5)
End If


NewChipFlag = 0
If OldChipName <> ChipName Then
    FileCopy App.Path & "\AlcorMP_698x2\New_INI\" & ChipName & "\ROM\ROM.Hex", App.Path & "\AlcorMP_698x2\ROM.Hex"
    FileCopy App.Path & "\AlcorMP_698x2\New_INI\" & ChipName & "\RAM\\RAM.Bin", App.Path & "\AlcorMP_698x2\RAM.Bin"
    FileCopy App.Path & "\AlcorMP_698x2\New_INI\" & ChipName & "\AlcorMP.ini", App.Path & "\AlcorMP_698x2\AlcorMP.ini"
    FileCopy App.Path & "\AlcorMP_698x2\New_INI\" & ChipName & "\PE.bin", App.Path & "\AlcorMP_698x2\PE.bin"
    FileCopy App.Path & "\AlcorMP_698x2\New_INI\" & ChipName & "\FT.ini", App.Path & "\FT.ini"
    FileCopy App.Path & "\AlcorMP_698x2\New_INI\" & ChipName & "\FT.ini", App.Path & "\AlcorMP_698x2\FT.ini"
    NewChipFlag = 1 ' force MP
End If
          
OldChipName = ChipName
 

'==============================================================
' when begin RW Test, must clear MP porgram
'===============================================================


'(1)  /// for Auto mode
winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
  Loop While winHwnd <> 0
End If

'(1)
winHwnd = FindWindow(vbNullString, AU6992MPCaption)
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, AU6992MPCaption)
  Loop While winHwnd <> 0
End If
'(2)
winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
  Do
  rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
  Call MsecDelay(0.5)
  winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
  Loop While winHwnd <> 0
End If
MPTester.Print "ContFail="; ContFail
MPTester.Print "MPContFail="; MPContFail
 
 
 '====================================
 '  Fix Card
 '====================================
' GoTo T1
 If (ContFail >= 5) Or (MPTester.Check1.Value = 1) Or (NewChipFlag = 1) Or (ForceMP_Flag = True) Then
 
 
   If MPTester.NoMP.Value = 1 Then
        
        If (NewChipFlag = 0) And (MPTester.Check1.Value = 0) Then  ' force condition
           GoTo RW_Test_Label
        End If
    End If
       
    
    If MPTester.ResetMPFailCounter.Value = 1 Then
       ContFail = 0
    End If
    
 '==============================================================
' when begin MP, must close RW porgram
'===============================================================
   MPFlag = 1
 
    winHwnd = FindWindow(vbNullString, "UFD Test")
    If winHwnd <> 0 Then
      Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "UFD Test")
      Loop While winHwnd <> 0
    End If
 
       '  power on
       cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        Call PowerSet(3)   ' close power to disable chip
       Call MsecDelay(0.5)  ' power for load MPDriver
       MPTester.Print "wait for MP Ready"
       Call LoadMP_Click_AU6992
 
        OldTimer = Timer
        AlcorMPMessage = 0
        Debug.Print "begin"
        Do
           ' DoEvents
            If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                AlcorMPMessage = mMsg.message
                TranslateMessage mMsg
                DispatchMessage mMsg
            End If
            PassTime = Timer - OldTimer
            'Debug.Print AlcorMPMessage
        Loop Until AlcorMPMessage = WM_FT_MP_START Or PassTime > 30 _
              Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
              
        
        MPTester.Print "Ready Time="; PassTime
        
        '====================================================
        '  handle MP load time out, the FAIL will be Bin3
        '====================================================
        If PassTime > 30 Then
          '(1)
           MPTester.Print "MP Ready Fail"
            winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, AU6992MPCaption1)
              Loop While winHwnd <> 0
            End If
           '(2)
            winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
            If winHwnd <> 0 Then
              Do
              rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
              Call MsecDelay(0.5)
              winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
              Loop While winHwnd <> 0
            End If
            
        
             MPTester.TestResultLab = "Bin3:MP Ready Fail"
             TestResult = "Bin3"
             MPTester.Print "MP Ready Fail"
     
              
            Exit Sub
        End If
        
        '====================================================
        '  MP begin
        '====================================================
        
        If AlcorMPMessage = WM_FT_MP_START Then
         
             
             cardresult = DO_WritePort(card, Channel_P1A, &HFD)  ' sel chip
              Call PowerSet(500)   ' close power to disable chip
             
             
            Dim TimerCounter As Integer
            Dim TmpString As String
            
             
            Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName("vid")
            Loop While TmpString = "" And TimerCounter < 150
             
            Call MsecDelay(0.3)
             
             If TmpString = "" Then   ' can not find device after 15 s
             
               TestResult = "Bin2"
               MPTester.TestResultLab = "Bin2:MP UNKNOW Fail when enter MP"
               Exit Sub
             End If
             
             Call MsecDelay(2.5)
               
             MPTester.Print " MP Begin....."
             
             Call StartMP_Click_AU6992
   
              ReMP_Flag = 0
              OldTimer = Timer
              AlcorMPMessage = 0
                
                Do
                   'DoEvents
                    If PeekMessage(mMsg, 0, 0, 0, PM_REMOVE) Then
                        AlcorMPMessage = mMsg.message
                        TranslateMessage mMsg
                        DispatchMessage mMsg
                            
                        If (AlcorMPMessage = WM_FT_MP_FAIL) And (ReMP_Flag = 0) Then
                            'ReMP_Flag = 1
                            AlcorMPMessage = 1
                            cardresult = DO_WritePort(card, Channel_P1A, &HFF)  'close power
                            Call MsecDelay(0.3)
                            cardresult = DO_WritePort(card, Channel_P1A, &HFD)  'open power
                            Call MsecDelay(2.2)
                            Call RefreshMP_Click_AU6992
                            Call MsecDelay(0.5)
                            Call StartMP_Click_AU6992
                            
                            ReMP_Counter = ReMP_Counter + 1
                            If ReMP_Counter >= ReMP_Limit Then
                                ReMP_Flag = 1
                                ReMP_Counter = 0
                            End If
                        End If
                    End If
                    
                    PassTime = Timer - OldTimer
                
                Loop Until AlcorMPMessage = WM_FT_MP_PASS _
                Or AlcorMPMessage = WM_FT_MP_FAIL _
                Or AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL _
                Or PassTime > 65 _
                Or AlcorMPMessage = WM_CLOSE Or AlcorMPMessage = WM_DESTROY
                
                If AlcorMPMessage = WM_FT_MP_PASS Then
                    If ReMP_Flag = 0 Then
                        MsecDelay (MPIdleTime * (ReMP_Limit - ReMP_Counter))
                    End If
                    ReMP_Counter = 0
                End If

                MPTester.Print "MP work time="; PassTime
                 MPTester.MPText.Text = Hex(AlcorMPMessage)
                '================================================
                '  Handle MP work time out error
                '===============================================
                
               ' time out fail
                If PassTime > 65 Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Time out Fail"
                    MPTester.Print "MP Time out Fail"
                    '(1)
                    winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                      Loop While winHwnd <> 0
                    End If
                    
                    '(2)
                      winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    Exit Sub
                End If
                
                ' MP fail
                If AlcorMPMessage = WM_FT_MP_FAIL Then
                    MPContFail = MPContFail + 1
                    TestResult = "Bin3"
                    MPTester.TestResultLab = "Bin3:MP Function Fail"
                    MPTester.Print "MP Function Fail"
                    
                    winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                    If winHwnd <> 0 Then
                      Do
                      rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                      Call MsecDelay(0.5)
                      winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                      Loop While winHwnd <> 0
                    End If
                    
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                        If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                        End If
                    
                    Exit Sub
                End If
                
                
                 'unknow fail
                 If AlcorMPMessage = WM_FT_MP_UNKNOW_FAIL Then
                     MPContFail = MPContFail + 1
                    TestResult = "Bin2"
                    MPTester.TestResultLab = "Bin2:MP UNKNOW Fail"
                    MPTester.Print "MP UNKNOW Fail"
                    
                   winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                    If winHwnd <> 0 Then
                     Do
                     rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                     Call MsecDelay(0.5)
                     winHwnd = FindWindow(vbNullString, AU6992MPCaption)
                     Loop While winHwnd <> 0
                     End If
                     
                     winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
                     If winHwnd <> 0 Then
                          Do
                          rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                          Call MsecDelay(0.5)
                          winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
                          Loop While winHwnd <> 0
                     End If
                     
                     
                     
                    Exit Sub
                End If
                 
                
                ' mp pass
                If AlcorMPMessage = WM_FT_MP_PASS Then
                     MPTester.TestResultLab = "MP PASS"
                    MPContFail = 0
                    MPTester.Print "MP PASS"
                End If
        End If
   
End If
'=========================================
 '    Close MP program
 '=========================================
winHwnd = FindWindow(vbNullString, AU6992MPCaption)
If winHwnd <> 0 Then
  Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, AU6992MPCaption)
  Loop While winHwnd <> 0
    
    Call MsecDelay(0.2)
    KillProcess ("AlcorMP.exe")

End If
    
 winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
 
If winHwnd <> 0 Then
    Do
        rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        Call MsecDelay(0.5)
        winHwnd = FindWindow(vbNullString, "重新啟動量產程式")
    Loop While winHwnd <> 0
End If

  Dim pid As Long          ' unload driver
  Dim hProcess As Long
  Dim ExitEvent As Long
 
  pid = Shell(App.Path & "\AlcorMP_698x2\loaddrv.exe uninstall_058F6387")
  hProcess = OpenProcess(SYNCHRONIZE + PROCESS_QUERY_INFORMATION + PROCESS_TERMINATE, 0, pid)
  ExitEvent = WaitForSingleObject(hProcess, INFINITE)
  Call CloseHandle(hProcess)
  KillProcess ("LoadDrv.exe")
 
 
                        
 '=========================================
 '    POWER on
 '=========================================
'T1:
RW_Test_Label:
 
 If MPFlag = 1 Then
         Call PowerSet(3)
          cardresult = DO_WritePort(card, Channel_P1A, &HFF)
        
         Call MsecDelay(0.5)  ' power of to unload MPDriver

           cardresult = DO_WritePort(card, Channel_P1A, &HFA)  'sel socket
           Call PowerSet(1500)
     
        
         Call MsecDelay(1.2)
        MPFlag = 0
 Else
          cardresult = DO_WritePort(card, Channel_P1A, &HFA)
         Call PowerSet(1500)
         
         Call MsecDelay(1.2)
End If
         Call LoadRWTest_Click_AU6992

        
        
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
     '   GoTo T2
       If PassTime > 5 Then
           TestResult = "Bin3"
           MPTester.TestResultLab = "Bin3:RW Ready Fail"
          
            winHwnd = FindWindow(vbNullString, "UFD Test")
            If winHwnd <> 0 Then
              Do
                rt2 = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
                Call MsecDelay(0.5)
                winHwnd = FindWindow(vbNullString, "UFD Test")
              Loop While winHwnd <> 0
            End If
       
            Exit Sub
       End If
         
T2:
        OldTimer = Timer
        AlcorMPMessage = 0
        MPTester.Print "RW Tester begin test........"
        Call StartRWTest_Click_AU6988
        
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
            Or AlcorMPMessage = WM_FT_CHECK_CERBGPO_FAIL _
            Or AlcorMPMessage = WM_FT_CHECK_HW_CODE_FAIL _
            Or AlcorMPMessage = WM_FT_PHYREAD_FAIL _
            Or AlcorMPMessage = WM_FT_ECC_FAIL _
            Or AlcorMPMessage = WM_FT_NOFREEBLOCK_FAIL _
            Or AlcorMPMessage = WM_FT_LODECODE_FAIL _
            Or AlcorMPMessage = WM_FT_RELOADCODE_FAIL _
            Or AlcorMPMessage = WM_FT_TESTUNITREADY_FAIL _
            Or PassTime > 10 _
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
                cardresult = DO_WritePort(card, Channel_P1A, &HFA)  ' power off
                Exit Sub
            End If
        
        End If
        
        
        Select Case AlcorMPMessage
        
        Case WM_FT_RW_UNKNOW_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:UnKnow Fail"
             ContFail = ContFail + 1
        
        Case WM_FT_CHECK_HW_CODE_FAIL
             TestResult = "Bin5"
             MPTester.TestResultLab = "Bin5:HW-ID Fail"
             ContFail = ContFail + 1
        
        Case WM_FT_TESTUNITREADY_FAIL
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:TestUnitReady Fail"
             ContFail = ContFail + 1

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
        
        Case WM_FT_RELOADCODE_FAIL
              TestResult = "Bin4"
              MPTester.TestResultLab = "Bin4:ReLoadCode FAIL "
              ContFail = ContFail + 1
        
        Case WM_FT_ECC_FAIL
              TestResult = "Bin5"
              MPTester.TestResultLab = "Bin5:ECC FAIL "
              ContFail = ContFail + 1
                    
        Case WM_FT_RW_RW_PASS
               
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
              
               End If
               
        Case Else
             TestResult = "Bin2"
             MPTester.TestResultLab = "Bin2:Undefine Fail"
          
             ContFail = ContFail + 1

        End Select
                               
       cardresult = DO_WritePort(card, Channel_P1A, &HFA)
         Call PowerSet(1500)
              
End Sub

