Attribute VB_Name = "AU6376MDL"
Option Explicit
Public Sub AU6376TestSub()

If ChipName = "AU6376BLF20" Then

   Call AU6376BLF20TestSub
   
 ElseIf ChipName = "AU6376BLF22" Then
   
   Call AU6376BLF22TestSub
  ElseIf ChipName = "AU6376FLF22" Then
   
   Call AU6376FLF22TestSub
   
 End If


End Sub

Public Sub AU6376TestAllSub()

 If ChipName = "AU6376KLF20" Then
                    ChipName = "AU6376JLF20"
                    Call MultiSlotTest
                   End If
                  
                  
                   
                    If ChipName = "AU6376ALF20" Then
                    ChipName = "AU6376"
                   Call MultiSlotTestAU6376ALF20
                   End If
                   
                   
                  If ChipName = "AU6376ALF21" Then
                    ChipName = "AU6376"
                   Call MultiSlotTestAU6376ALF21
                   End If
                   
                   
                   If ChipName = "AU6376ALF22" Or ChipName = "AU6376ELF22" Then
                    ChipName = "AU6376"
                   Call MultiSlotTestAU6376ALF22
                   End If
                   
                   
                   If ChipName = "AU6376ALO10" Then
                      ChipName = "AU6376"
                      Call MultiSlotTestAU6376ALO10
                    
                   
                    ElseIf ChipName = "AU6376ALO11" Then
                      ChipName = "AU6376"
                      Call MultiSlotTestAU6376ALO11
                  
                        ElseIf ChipName = "AU6376ALO14" Then
                      ChipName = "AU6376"
                      Call MultiSlotTestAU6376ALO14
                   
                     ElseIf ChipName = "AU6376ALO12" Then
                      ChipName = "AU6376"
                      Call MultiSlotTestAU6376ALO12
                  
                      ElseIf ChipName = "AU6376ALO13" Then
                        ChipName = "AU6376"
                      Call MultiSlotTestAU6376ALO13
                      ElseIf ChipName = "AU6376ALOT1" Then
                        ChipName = "AU6376"
                      Call MultiSlotTestAU6376ALOT1
                      ElseIf ChipName = "AU6376ALOT2" Then
                        ChipName = "AU6376"
                      Call MultiSlotTestAU6376ALOT2
                       ElseIf ChipName = "AU6376ALOT3" Then
                        ChipName = "AU6376"
                      Call MultiSlotTestAU6376ALOT3
                   End If

End Sub
  
Public Sub AU6376FLF20TestSub()

  If ChipName <> "AU6333NC" And ChipName <> "AU6331NC" Then
                    If PCI7248InitFinish = 0 Then
                      PCI7248Exist
                    End If
                    If ChipName = "AU6333A1" Or ChipName = "AU6333BL" Then
                     CardResult = DO_WritePort(card, Channel_P1A, &H0)
                     Call MsecDelay(0.2)
                    CardResult = DO_WritePort(card, Channel_P1A, &H6)
                    Call MsecDelay(1#)
                    End If
                    
                    If ChipName = "AU6333EL" Then
                     CardResult = DO_WritePort(card, Channel_P1A, &H9)
                     Call MsecDelay(0.2)
                    CardResult = DO_WritePort(card, Channel_P1A, &HE)
                    Call MsecDelay(1#)
                    End If
                    
                     If ChipName = "AU6376FLF20" Then
                     result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
                     CardResult = DO_WritePort(card, Channel_P1B, &H0)
                     Call MsecDelay(0.2)
                     result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                      
                    
                     CardResult = DO_WritePort(card, Channel_P1A, &H3E)
                    Call MsecDelay(1#)
                    End If
                    
                 End If
            
                  
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                rv5 = 0
                rv6 = 0
                rv7 = 0
                
                LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                
                Tester.Print LBA
                
             
                
                If ChipName = "AU6376FLF20" Then
                
                    ClosePipe
                    rv0 = CBWTest_New_no_card(1, 1, "vid_058f")
                    Call LabelMenu(0, rv0, 1)
                    ClosePipe
                    rv1 = CBWTest_New_no_card(0, rv0, "vid_058f")
                    Call LabelMenu(1, rv1, rv0)
                    ClosePipe
                    
                Else
                      ClosePipe
                    rv0 = CBWTest_New_no_card(0, 1, "vid_058f")
                    'Print "a1"
                    Call LabelMenu(0, rv0, 1)
                    ClosePipe
                    rv1 = CBWTest_New_no_card(1, rv0, "vid_058f")
                      Call LabelMenu(3, rv1, rv0)
                     ClosePipe
                
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
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv1 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                      
                    TestResult = "MS_RF"
                ElseIf rv1 * rv0 = PASS Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                End If
                
                Tester.Print "Test Result"; TestResult
                       
               If ChipName <> "AU6333NC" And ChipName <> "AU6331NC" Then
                    TestResult = "PASS"
               Else
                    Tester.MSComm1.OutBufferCount = 0
                    Tester.MSComm1.Output = TestResult   ' send out test result
                  
               End If
            
                 
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 card change bit fail"
                If ChipName = "AU6376FLF20" Then
                   Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 card change bit fail"
                Else
                   Tester.Print rv1, " \\MS :0 Unknow device, 1 pass ,2 card change bit fail"
                End If
                If TestResult = "PASS" Then
                     TestResult = ""
                     
                     If ChipName = "AU6333NC" Or ChipName = "AU6331NC" Then
                        ChipName = ""
                        Do
                           Call MsecDelay(0.1)
                              DoEvents
                              buf = Tester.MSComm1.Input
                              ChipName = ChipName & buf
                     
                        Loop Until InStr(1, ChipName, "PASS") <> 0
                        
                     Else
                     
                       If ChipName = "AU6376FLF20" Then
                        CardResult = DO_WritePort(card, Channel_P1A, &H0)
                        Call MsecDelay(0.1)
                       Else
                        CardResult = DO_WritePort(card, Channel_P1A, &H0)
                        Call MsecDelay(1#)
                      End If
                        
           
                     End If
                 
                 '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                 '
                 '  R/W test
                 '
                 '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                 
                
                'initial return value
                
                
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                
                Tester.Label3.BackColor = RGB(255, 255, 255)
                 Tester.Label4.BackColor = RGB(255, 255, 255)
                 Tester.Label5.BackColor = RGB(255, 255, 255)
                 Tester.Label6.BackColor = RGB(255, 255, 255)
                 Tester.Label7.BackColor = RGB(255, 255, 255)
                 Tester.Label8.BackColor = RGB(255, 255, 255)
                
                
                If ChipName = "AU6376FLF20" Then
                
                    ClosePipe
                    rv0 = CBWTest_New(1, 1, "vid_058f")
                    Call LabelMenu(0, rv0, 1)
                    ClosePipe
                    rv1 = CBWTest_New(0, rv0, "vid_058f")
                    Call LabelMenu(1, rv1, rv0)
                    ClosePipe
                    
                Else
                
                    ClosePipe
                    rv0 = CBWTest_New(0, 1, "vid_058f")
                    Call LabelMenu(0, rv0, 1)
                    ClosePipe
                    rv1 = CBWTest_New(1, rv0, "vid_058f")
                    Call LabelMenu(3, rv1, rv0)
                    ClosePipe
                
                End If
               
                 If rv1 = 1 Then
                                CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                  
                                  
                                 If ChipName <> "AU6376FLF20" Then
                                    If LightOff <> 240 Then
                                      UsbSpeedTestResult = GPO_FAIL
                                       rv1 = 2
                                    End If
                                    
                                 Else
                                      If LightOff <> 252 Then
                                      UsbSpeedTestResult = GPO_FAIL
                                       rv1 = 2
                                    End If
                                End If
                                 
                 End If
                
                
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                If ChipName = "AU6376FLF20" Then
                 Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Else
                 Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                End If
                 Tester.Print "LBA="; LBA
                
                
                'If rv0 = 1 And rv1 = 1 And rv2 = 1 And rv3 = 1 Then
                   ' TestResult = "PASS"
                'End If
                
                        
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
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv1 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                        End If
                
               ' If rv2 * rv3 * rv1 = 0 Then
                '    MsgBox "rturn error"
                'End If
                  
                End If


End Sub

Public Sub AU6376FLF22TestSub()

  
                    If PCI7248InitFinish = 0 Then
                      PCI7248Exist
                    End If
                   
                    
                   
                  
                    
                    
                     result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
                     CardResult = DO_WritePort(card, Channel_P1B, &H0)
                     Call MsecDelay(0.2)
                     result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                     CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                    Call MsecDelay(1#)
                    
            
                  
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                rv5 = 0
                rv6 = 0
                rv7 = 0
                
                LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                
                Tester.Print LBA
                
             
                
               
                
                    ClosePipe
                     rv0 = CBWTest_New_no_card(1, 1, "vid_058f")
                     Call LabelMenu(0, rv0, 1)
                    ClosePipe
                    
                    rv1 = CBWTest_New_no_card(0, rv0, "vid_058f")
                    If rv1 = 1 Then
                       rv1 = SetOverCurrent(rv1)
                       If rv1 = 0 Then
                          rv1 = 2
                       End If
                     End If
                       
                    Call LabelMenu(1, rv1, rv0)
                    ClosePipe
              
                
                
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
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv1 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                      
                    TestResult = "MS_RF"
                ElseIf rv1 * rv0 = PASS Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                End If
                
                Tester.Print "Test Result"; TestResult
                       
             
                 
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 card change bit fail"
               
                   Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 card change bit fail"
                
                If TestResult = "PASS" Then
                     TestResult = ""
                     
                    
                   
                     
                      
                        CardResult = DO_WritePort(card, Channel_P1A, &H0)
                        Call MsecDelay(0.2)
                       
                 
                 '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                 '
                 '  R/W test
                 '
                 '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                 
                
                'initial return value
                
                
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                
                Tester.Label3.BackColor = RGB(255, 255, 255)
                 Tester.Label4.BackColor = RGB(255, 255, 255)
                 Tester.Label5.BackColor = RGB(255, 255, 255)
                 Tester.Label6.BackColor = RGB(255, 255, 255)
                 Tester.Label7.BackColor = RGB(255, 255, 255)
                 Tester.Label8.BackColor = RGB(255, 255, 255)
                
                
               
                    ClosePipe
                    rv0 = CBWTest_New(1, 1, "vid_058f")
                    Call LabelMenu(0, rv0, 1)
                    ClosePipe
                    rv1 = CBWTest_New(0, rv0, "vid_058f")
                    
                    If rv1 = 1 Then
                       rv1 = Read_OverCurrent(0, 0, 64)
                       If rv1 = 0 Then
                        rv1 = 2
                       End If
                    End If
                       
                       
                    Call LabelMenu(1, rv1, rv0)
                    ClosePipe
                                CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                  
                                  
                                
                                    If LightOff <> 252 Then
                                      UsbSpeedTestResult = GPO_FAIL
                                       rv1 = 2
                                    End If
                                    
         
                
                
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                 Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
              
                 Tester.Print "LBA="; LBA
                
                
                'If rv0 = 1 And rv1 = 1 And rv2 = 1 And rv3 = 1 Then
                   ' TestResult = "PASS"
                'End If
                
                        
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
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv1 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                        End If
                
             
                End If


End Sub
Public Sub AU6376BLF22TestSub()
  LBA = LBA + 1
                TestResult = ""
                
                
                
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                 
                
                    CardResult = DO_WritePort(card, Channel_P1A, &HE)   ' 0111 1111
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                 
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                '  R/W test
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                
                'initial return value
                
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
                
                   
                   CardResult = DO_WritePort(card, Channel_P1A, &HE)   ' 0111 1111
                ClosePipe
                 
                 rv0 = CBWTest_New_no_card(0, 1, "vid_058f")
                 
                 If rv0 = 1 Then
                    rv0 = SetOverCurrent(rv0)
                    If rv0 <> 1 Then
                      rv0 = 2
                    End If
                 Else
                    GoTo AU6376TestResult
                 End If
                ' Tester.print "a4"
                ClosePipe
                 
                
                  CardResult = DO_WritePort(card, Channel_P1A, &HA)   ' 0111 1111
                  
                  Call MsecDelay(0.2)
                ClosePipe
                    rv0 = CBWTest_New(0, rv0, "vid_058f")
                    If rv0 = 1 Then
                        rv0 = Read_OverCurrent(0, 0, 64)
                        If rv0 <> 1 Then
                          rv0 = 2
                        End If
                        
                    Else
                        GoTo AU6376TestResult
                    End If
                    Call LabelMenu(1, rv0, 1)
                ClosePipe
                
                  
                   
                    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 If LightOff <> 252 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                  Else
                                  rv4 = 1
                                 End If
                     Call LabelMenu(3, rv4, rv0)
                  
                 
                
 
                Tester.Print rv0, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv4, " \\CF :0 Unknow device, 1 pass ,2 GPO FAIL, 4 preious slot fail"
               Tester.Print " <<<<<<<<<<<<<<<<<<------------------>>>>>>>>>>>>>>>>>>>>>>"
                Tester.Print "LBA="; LBA
                
AU6376TestResult:     If rv0 = UNKNOW Then
                    UnknowDeviceFail = UnknowDeviceFail + 1
                    TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv0 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
                    TestResult = "CF_RF"
                     TestResult = "UNKNOW"
                ElseIf rv4 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv4 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                ElseIf rv0 * rv4 = PASS Then
                    TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                End If
                
End Sub

Public Sub AU6376BLF20TestSub()
  LBA = LBA + 1
                TestResult = ""
                
                
                
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                 
                 If ChipName = "AU6376BLF20" Then
                    CardResult = DO_WritePort(card, Channel_P1A, &HA)   ' 0111 1111
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                End If
                
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                '  R/W test
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                
                'initial return value
                
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
                ClosePipe
                    rv0 = CBWTest_New(0, 1, "vid_058f")
                    Call LabelMenu(1, rv0, 1)
                ClosePipe
                
                  If ChipName = "AU6376BLF20" Then
                   
                    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 If LightOff <> 252 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                  Else
                                  rv4 = 1
                                 End If
                     Call LabelMenu(3, rv4, rv0)
                  
                  End If
                
                Tester.Print rv0, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv4, " \\CF :0 Unknow device, 1 pass ,2 GPO FAIL, 4 preious slot fail"
               Tester.Print " <<<<<<<<<<<<<<<<<<------------------>>>>>>>>>>>>>>>>>>>>>>"
                Tester.Print "LBA="; LBA
                
                If rv0 = UNKNOW Then
                    UnknowDeviceFail = UnknowDeviceFail + 1
                    TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv0 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
                    TestResult = "CF_RF"
                     TestResult = "UNKNOW"
                ElseIf rv4 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv4 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                ElseIf rv0 * rv4 = PASS Then
                    TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                End If
                
End Sub
