Attribute VB_Name = "AU6471MDL"
Public Sub AU6471F21TestSub()

If ChipName = "AU6471GLF21" Then

  Call AU6471GLF21TestSub

End If

If ChipName = "AU6471FLF21" Then

  Call AU6471FLF21TestSub

End If




End Sub

Public Sub AU6471F22TestSub()

If ChipName = "AU6471GLF22" Then

  Call AU6471GLF22TestSub

End If

If ChipName = "AU6471FLF22" Then

  Call AU6471FLF22TestSub

End If

If ChipName = "AU6471FLF23" Then

  Call AU6471FLF23TestSub

End If

If ChipName = "AU6471GLF23" Then

  Call AU6471GLF23TestSub

End If

If ChipName = "AU6471GLF24" Then

  Call AU6471GLF24TestSub

End If

If ChipName = "AU6471GLF04" Then

  Call AU6471GLF04TestSub

End If

If ChipName = "AU6471JLS10" Then

  Call AU6471JLS10SortingSub

End If




End Sub
Public Sub AU6471FLF21TestSub()

Tester.Print "AU6471FL is nb mode"
Dim i As Integer
       
               If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
               
               ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
               '  CardResult = DO_WritePort(card, Channel_P1B, &H0)
                
                LBA = LBA + 1
                         
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                rv5 = 0
                rv6 = 0
                rv7 = 0
             
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                    If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 Call MsecDelay(0.05)
              
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(0.8)    'power on time
              
                '===============================================
                '  SD Card test
                '================================================
           
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                     Dim ChipString  As String
                       ChipString = "vid_1984"
                       If GetDeviceName(ChipString) <> "" Then
                    Tester.Print "NB mode Test Fail"
                    TestResult = "Bin2"
                    Call LabelMenu(0, 2, 1)
                                 
                    Exit Sub
                     End If
                
             
                   Call MsecDelay(0.01)
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(0.8)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                     
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
                      
                      
                       For i = 1 To 20
                      
                        If rv0 = 1 Then
                           
                            ClosePipe
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                                
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                              
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                           
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                              
                            ClosePipe
                        End If
              
                        If rv0 <> 1 Then
                        GoTo AU6371ELResult
                        End If
                           
                        Next
                     
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                                    
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
              
                  CardResult = DO_WritePort(card, Channel_P1A, &H7C) ' SD + CF
                  
                    Call MsecDelay(0.1)
                 If rv0 = 1 Then
                     CardResult = DO_WritePort(card, Channel_P1A, &H7D)
                 End If
                    Call MsecDelay(0.1)
                OpenPipe
                rv1 = ReInitial(0)
               
                 ClosePipe
                 rv1 = CBWTest_New(0, rv0, ChipString)
                 ClosePipe
  
                  
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
                   
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
   
                 
                 Call LabelMenu(1, rv1, rv0)
            
                      Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
                   CardResult = DO_WritePort(card, Channel_P1A, &H79)  '1001
                  
               
                  
                  Call MsecDelay(0.1)
              
                     CardResult = DO_WritePort(card, Channel_P1A, &H7B) '1011
                
               
                    Call MsecDelay(0.1)
                    
                   OpenPipe
                rv2 = ReInitial(0)
                  
                 ClosePipe
                 rv2 = CBWTest_New(0, rv1, ChipString)
                 Call LabelMenu(21, rv2, rv1)
                
                     Tester.Print rv2, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
              
               
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H73)  '0011
                  
                    Call MsecDelay(0.1)
                 
                 
                   CardResult = DO_WritePort(card, Channel_P1A, &H77) '0111
                
               
                    Call MsecDelay(0.1)
                    
                   OpenPipe
                rv3 = ReInitial(0)
                
                ClosePipe
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  MS Card test
                '================================================
                  
                     rv4 = rv3
                
               
                   '  Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)  '0011
                  
                    Call MsecDelay(0.1)
                 
                 
                   CardResult = DO_WritePort(card, Channel_P1A, &H5F) '0111
                
               
                    Call MsecDelay(0.1)
                    
                   OpenPipe
                rv5 = ReInitial(0)
                   
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                
                
                  CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
AU6371ELResult:
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
                        ElseIf rv2 = WRITE_FAIL Or rv3 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv2 = READ_FAIL Or rv3 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                         ElseIf rv4 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Or rv5 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                       
                            
                        ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub
Public Sub AU6471FLF22TestSub()

Tester.Print "AU6471FL is nb mode"
Dim i As Integer
       
               If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
               
               ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
               '  CardResult = DO_WritePort(card, Channel_P1B, &H0)
                
                LBA = LBA + 1
                         
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                rv5 = 0
                rv6 = 0
                rv7 = 0
             
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                    If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 Call MsecDelay(0.05)
              
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(0.8)    'power on time
              
                '===============================================
                '  SD Card test
                '================================================
           
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                     Dim ChipString  As String
                       ChipString = "vid_1984"
                       If GetDeviceName(ChipString) <> "" Then
                    Tester.Print "NB mode Test Fail"
                    TestResult = "Bin2"
                    Call LabelMenu(0, 2, 1)
                                 
                    Exit Sub
                     End If
                
             
                   Call MsecDelay(0.01)
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(0.8)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                     
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      If rv0 = 1 Then
                          rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                          If rv0 <> 1 Then
                          
                              rv0 = 2
                              Tester.Print "SD bus width Fail"
                          End If
                      End If
                          
                      
                      
                      ClosePipe
                      
                      
                       For i = 1 To 20
                      
                        If rv0 = 1 Then
                           
                            ClosePipe
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                                
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                              
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                           
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                              
                            ClosePipe
                        End If
              
                        If rv0 <> 1 Then
                        GoTo AU6371ELResult
                        End If
                           
                        Next
                     
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                                    
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
              
                  CardResult = DO_WritePort(card, Channel_P1A, &H7C) ' SD + CF
                  
                    Call MsecDelay(0.1)
                 If rv0 = 1 Then
                     CardResult = DO_WritePort(card, Channel_P1A, &H7D)
                 End If
                    Call MsecDelay(0.1)
                OpenPipe
                rv1 = ReInitial(0)
               
                 ClosePipe
                 rv1 = CBWTest_New(0, rv0, ChipString)
                 ClosePipe
  
                  
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
                   
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
   
                 
                 Call LabelMenu(1, rv1, rv0)
            
                      Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
                   CardResult = DO_WritePort(card, Channel_P1A, &H79)  '1001
                  
               
                  
                  Call MsecDelay(0.1)
              
                     CardResult = DO_WritePort(card, Channel_P1A, &H7B) '1011
                
               
                    Call MsecDelay(0.1)
                    
                   OpenPipe
                rv2 = ReInitial(0)
                  
                 ClosePipe
                 rv2 = CBWTest_New(0, rv1, ChipString)
                 Call LabelMenu(21, rv2, rv1)
                
                     Tester.Print rv2, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
              
               
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H73)  '0011
                  
                    Call MsecDelay(0.1)
                 
                 
                   CardResult = DO_WritePort(card, Channel_P1A, &H77) '0111
                
               
                    Call MsecDelay(0.1)
                    
                   OpenPipe
                rv3 = ReInitial(0)
                
                ClosePipe
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  MS Card test
                '================================================
                  
                     rv4 = rv3
                
               
                   '  Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)  '0011
                  
                    Call MsecDelay(0.1)
                 
                 
                   CardResult = DO_WritePort(card, Channel_P1A, &H5F) '0111
                
               
                    Call MsecDelay(0.1)
                    
                   OpenPipe
                rv5 = ReInitial(0)
                   
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                 If rv5 = 1 Then
                          rv5 = Read_MS_Speed_AU6471(0, 0, 64, "4Bits")
                          If rv5 <> 1 Then
                          
                              rv5 = 2
                              Tester.Print "SD bus width Fail"
                          End If
                      End If
               
                      
                   
                
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                
                
                  CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
AU6371ELResult:
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
                        ElseIf rv2 = WRITE_FAIL Or rv3 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv2 = READ_FAIL Or rv3 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                         ElseIf rv4 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Or rv5 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                       
                            
                        ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub
Public Sub AU6471FLF23TestSub()

Tester.Print "AU6471FL is nb mode"
Dim i As Integer
       
               If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
               
               ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
               '  CardResult = DO_WritePort(card, Channel_P1B, &H0)
                
                LBA = LBA + 1
                         
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                rv5 = 0
                rv6 = 0
                rv7 = 0
             
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                    If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 Call MsecDelay(0.05)
              
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(0.8)    'power on time
              
                '===============================================
                '  SD Card test
                '================================================
           
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                     Dim ChipString  As String
                       ChipString = "vid_058f"
                       If GetDeviceName(ChipString) <> "" Then
                    Tester.Print "NB mode Test Fail"
                    TestResult = "Bin2"
                    Call LabelMenu(0, 2, 1)
                                 
                    Exit Sub
                     End If
                
             
                   Call MsecDelay(0.01)
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(0.8)
                    
                     
                     
                             
                           
                      ClosePipe
                     
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      If rv0 = 1 Then
                          rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                          If rv0 <> 1 Then
                          
                              rv0 = 2
                              Tester.Print "SD bus width Fail"
                          End If
                      End If
                          
                      
                      
                      ClosePipe
                      
                      
                       For i = 1 To 20
                      
                        If rv0 = 1 Then
                           
                            ClosePipe
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                                
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                              
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                           
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                              
                            ClosePipe
                        End If
              
                        If rv0 <> 1 Then
                        GoTo AU6371ELResult
                        End If
                           
                        Next
                     
                        
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                                    
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
              
                  CardResult = DO_WritePort(card, Channel_P1A, &H7C) ' SD + CF
                  
                    Call MsecDelay(0.1)
                 If rv0 = 1 Then
                     CardResult = DO_WritePort(card, Channel_P1A, &H7D)
                 End If
                    Call MsecDelay(0.1)
                OpenPipe
                rv1 = ReInitial(0)
               
                 ClosePipe
                 rv1 = CBWTest_New(0, rv0, ChipString)
                 ClosePipe
  
                  
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
                   
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
   
                 
                 Call LabelMenu(1, rv1, rv0)
            
                      Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
                   CardResult = DO_WritePort(card, Channel_P1A, &H79)  '1001
                  
               
                  
                  Call MsecDelay(0.1)
              
                     CardResult = DO_WritePort(card, Channel_P1A, &H7B) '1011
                
               
                    Call MsecDelay(0.1)
                    
                   OpenPipe
                rv2 = ReInitial(0)
                  
                 ClosePipe
                 rv2 = CBWTest_New(0, rv1, ChipString)
                 Call LabelMenu(21, rv2, rv1)
                
                     Tester.Print rv2, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
              
               
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H73)  '0011
                  
                    Call MsecDelay(0.1)
                 
                 
                   CardResult = DO_WritePort(card, Channel_P1A, &H77) '0111
                
               
                    Call MsecDelay(0.1)
                    
                   OpenPipe
                rv3 = ReInitial(0)
                
                ClosePipe
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  MS Card test
                '================================================
                  
                     rv4 = rv3
                
               
                   '  Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)  '0011
                  
                    Call MsecDelay(0.1)
                 
                 
                   CardResult = DO_WritePort(card, Channel_P1A, &H5F) '0111
                
               
                    Call MsecDelay(0.1)
                    
                   OpenPipe
                rv5 = ReInitial(0)
                   
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                 If rv5 = 1 Then
                          rv5 = Read_MS_Speed_AU6471(0, 0, 64, "4Bits")
                          If rv5 <> 1 Then
                          
                              rv5 = 2
                              Tester.Print "SD bus width Fail"
                          End If
                      End If
               
                      
                   
                
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                
                
                  CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
AU6371ELResult:
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
                        ElseIf rv2 = WRITE_FAIL Or rv3 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv2 = READ_FAIL Or rv3 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                         ElseIf rv4 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Or rv5 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                       
                            
                        ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub

Public Sub AU6471FLTest20()
     ' AU6371
         ' AF, CF : common board
         ' DL, EL, JL : common board
         ' GL,HL : common board
         ' EL,HL : card to open power

          'Combo1.AddItem "AU6371AFF25" ' AU6371S3/20061213
          'Combo1.AddItem "AU6371BLF25"   'AU6371
          'Combo1.AddItem "AU6371CFF25"  'AU6371CF/20061221
          'Combo1.AddItem "AU6371CLF25"  'AU6371EL
          'Combo1.AddItem "AU6371DLF25"  'AU6371/20060822,AU6371DLF21
          'Combo1.AddItem "AU6371DFF25"  'AU6371DFT10/20070216
          'Combo1.AddItem "AU6371ELF25"  'AU6371EL/20061124
          'Combo1.AddItem "AU6371FLF25"  'AU6371S3/20061113
          'Combo1.AddItem "AU6371GLF25"   'AU6371GL
          'Combo1.AddItem "AU6371HLF25"   'AU6371HL
          'Combo1.AddItem "AU6371JLF25"   'AU6371EL


                Dim ChipString As String
                Dim i As Integer
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
                If Left(ChipName, 10) = "AU6471FLF2" Then
                  ChipName = "AU6371EL"
                  OldChipName = "AU6371ELF2"
                End If
                 
                 If Left(ChipName, 10) = "AU6371HLF2" Then
                  ChipName = "AU6371EL"
                  OldChipName = "AU6371HLF2"
                End If
                
                If ChipName = "AU6371S3" Then
                  ChipString = "vid_8751"
                Else
                  ChipString = "vid_058f"
                End If
                
                If ChipName = "AU6371EL" Or Left(ChipName, 10) = "AU6371GLF2" Then
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    If ChipName = "AU6371EL" Then
                        AU6371EL_BootTime = 1.2
                        If OldChipName = "AU6371ELF2" Then
                         AU6371EL_BootTime = 0.5
                        End If
                        
                    Else
                        AU6371EL_BootTime = 0.02
                    End If
                 Else
                    AU6371EL_SD = 0
                    AU6371EL_CF = 0
                    AU6371EL_XD = 0
                    AU6371EL_MS = 0
                    AU6371EL_MSP = 0
                    AU6371EL_BootTime = 0
                End If
            
               If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
               
               ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
               '  CardResult = DO_WritePort(card, Channel_P1B, &H0)
                
                LBA = LBA + 1
                         
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                rv5 = 0
                rv6 = 0
                rv7 = 0
             
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                    If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 Call MsecDelay(0.05)
                 
                 
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F - AU6371EL_SD)
                  
                 Call MsecDelay(1.2 + AU6371EL_BootTime)  'power on time
              
                '===============================================
                '  SD Card test
                '================================================
             '   If Left(ChipName, 10) = "AU6371DLF2" Then
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                
                  
                  
                 If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect On Fail"
                    End
                 End If
                 
             '  End If
                   Call MsecDelay(0.01)
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(0.01)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      ChipString = "vid_1984"
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
                      
                      
                       For i = 1 To 20
                      
                        If rv0 = 1 Then
                           
                            ClosePipe
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                                
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                              
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                           
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                              
                            ClosePipe
                        End If
              
                        If rv0 <> 1 Then
                        GoTo AU6371ELResult
                        End If
                           
                        Next
                     If Left(ChipName, 10) = "AU6371DLF2" Then
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                                    
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                     ElseIf Left(ChipName, 7) = "AU6366C" Then
                          
                         If rv0 <> 0 Then
                          If LightOn <> 175 Or LightOff <> 255 Then
                                    
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                      
                     Else
                     
                          If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HBF Then
                                    
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                     End If
                     
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
                If ChipName <> "AU6371S3" Then
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                  If CardResult <> 0 Then
                    MsgBox "Set CF Card Detect On Fail"
                    End
                  End If
                  
                  
                    Call MsecDelay(0.01)
                 If rv0 = 1 Then
                     CardResult = DO_WritePort(card, Channel_P1A, &H7D)
                 End If
                 If CardResult <> 0 Then
                    MsgBox "Set CF Card Detect Down Fail"
                    End
                 End If
                  Call MsecDelay(AU6371EL_BootTime * 2)
                 If ChipName = "AU6371EL" Then
                   ReaderExist = 0
                 End If
               
                 ClosePipe
                 rv1 = CBWTest_New(0, rv0, ChipString)
                 ClosePipe
  
                  
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
                   
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                
                
                
                 
               Else
                  rv1 = 1  '----------- AU6371S3 dp not have CF slot
                 
               End If
                 
                 Call LabelMenu(1, rv1, rv0)
            
                      Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
                   CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 If CardResult <> 0 Then
                     MsgBox "Set SMC Card Detect On Fail"
                     End
                  End If
                  
                  Call MsecDelay(0.01)
                 If rv1 = 1 Then
                     CardResult = DO_WritePort(card, Channel_P1A, &H7B)
                 End If
               
                If CardResult <> 0 Then
                     MsgBox "Set SMC Card Detect Down Fail"
                     End
                  End If
                 
                 Call MsecDelay(1.2 + AU6371EL_BootTime)  'power on time
                   If ChipName = "AU6371EL" Then
                   ReaderExist = 0
                 End If
                 
                 ClosePipe
                 rv2 = CBWTest_New(0, rv1, ChipString)
                 Call LabelMenu(21, rv2, rv1)
                
                     Tester.Print rv2, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
              
               
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.01)
                If rv2 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                End If
                Call MsecDelay(AU6371EL_BootTime)
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect Down Fail"
                    End
                 End If
                 
                   If ChipName = "AU6371EL" Then
                   ReaderExist = 0
                 End If
                
                ClosePipe
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  MS Card test
                '================================================
                  
                     rv4 = 1
                
               
                     Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
                
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect On Fail"
                    End
                 End If
                
                 Call MsecDelay(0.03)
                If rv4 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                End If
                  Call MsecDelay(AU6371EL_BootTime * 2)
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect Down Fail"
                    End
                 End If
                 If ChipName = "AU6371EL" Then
                   ReaderExist = 0
                 End If
                   
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                
                
                  CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
AU6371ELResult:
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
                        ElseIf rv2 = WRITE_FAIL Or rv3 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv2 = READ_FAIL Or rv3 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                         ElseIf rv4 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Or rv5 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                       
                            
                        ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub


Public Sub AU6471GLF21TestSub()
Tester.Print "AU6471GL is noraml mode"
     ' AU6371
         ' AF, CF : common board
         ' DL, EL, JL : common board
         ' GL,HL : common board
         ' EL,HL : card to open power

          'Combo1.AddItem "AU6371AFF25" ' AU6371S3/20061213
          'Combo1.AddItem "AU6371BLF25"   'AU6371
          'Combo1.AddItem "AU6371CFF25"  'AU6371CF/20061221
          'Combo1.AddItem "AU6371CLF25"  'AU6371EL
          'Combo1.AddItem "AU6371DLF25"  'AU6371/20060822,AU6371DLF21
          'Combo1.AddItem "AU6371DFF25"  'AU6371DFT10/20070216
          'Combo1.AddItem "AU6371ELF25"  'AU6371EL/20061124
          'Combo1.AddItem "AU6371FLF25"  'AU6371S3/20061113
          'Combo1.AddItem "AU6371GLF25"   'AU6371GL
          'Combo1.AddItem "AU6371HLF25"   'AU6371HL
          'Combo1.AddItem "AU6371JLF25"   'AU6371EL


                Dim ChipString As String
                Dim i As Integer
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                
                
                
                ChipString = "vid_1984"
                
            
               If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
               
               ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
               '  CardResult = DO_WritePort(card, Channel_P1B, &H0)
                
                LBA = LBA + 1
                         
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                rv5 = 0
                rv6 = 0
                rv7 = 0
             
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                    If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 Call MsecDelay(0.05)
                 
                 
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1.2)   'power on time
              
                 '=========================================
                '    Noraml Mode Test
                '=========================================
                
                    If GetDeviceName(ChipString) = "" Then
                    Tester.Print "Normal mode Test Fail"
                    TestResult = "Bin2"
                    Call LabelMenu(0, 2, 1)
                                 
                    Exit Sub
                     End If
                
                
              
                '===============================================
                '  SD Card test
                '================================================
             '   If Left(ChipName, 10) = "AU6371DLF2" Then
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                
                  
                
                 
             '  End If
                   Call MsecDelay(0.01)
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(0.01)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
                      
                      
                       For i = 1 To 20
                      
                        If rv0 = 1 Then
                           
                            ClosePipe
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                             
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                               
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                                                  
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                               
                            ClosePipe
                        End If
              
                        If rv0 <> 1 Then
                        GoTo AU6371DLResult
                        End If
                           
                        Next
                      
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                                    
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                     
                     
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                  If CardResult <> 0 Then
                    MsgBox "Set CF Card Detect On Fail"
                    End
                  End If
                  
                  
                    Call MsecDelay(0.01)
                 
                     CardResult = DO_WritePort(card, Channel_P1A, &H7D)
                 
               
                  
                
               
                 ClosePipe
                 rv1 = CBWTest_New(0, rv0, ChipString)
                 ClosePipe
  
                  
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
                   
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
         
                
                 
                 Call LabelMenu(1, rv1, rv0)
            
                      Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
              
                  
                  Call MsecDelay(0.01)
                 If rv1 = 1 Then
                     CardResult = DO_WritePort(card, Channel_P1A, &H7B)
                 End If
               
                 If CardResult <> 0 Then
                     MsgBox "Set SMC Card Detect Down Fail"
                     End
                  End If
                 
                 ClosePipe
                rv2 = CBWTest_New(0, rv1, "vid_058f")
                 Call LabelMenu(21, rv2, rv1)
                
                      Tester.Print rv2, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
              
              
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.01)
                If rv2 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                End If
                
                 
                   
                ClosePipe
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  MS Card test
                '================================================
                  
                     rv4 = rv3
               
                    ' Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect On Fail"
                    End
                 End If
                
                 Call MsecDelay(0.03)
                If rv4 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                End If
                   
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
               
                
               
                
                  CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
AU6371DLResult:
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
                        ElseIf rv2 = WRITE_FAIL Or rv3 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv2 = READ_FAIL Or rv3 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                         ElseIf rv4 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Or rv5 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                       
                            
                        ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub

Public Sub AU6471GLF22TestSub()
Tester.Print "AU6471GL is noraml mode"
     ' AU6371
         ' AF, CF : common board
         ' DL, EL, JL : common board
         ' GL,HL : common board
         ' EL,HL : card to open power

          'Combo1.AddItem "AU6371AFF25" ' AU6371S3/20061213
          'Combo1.AddItem "AU6371BLF25"   'AU6371
          'Combo1.AddItem "AU6371CFF25"  'AU6371CF/20061221
          'Combo1.AddItem "AU6371CLF25"  'AU6371EL
          'Combo1.AddItem "AU6371DLF25"  'AU6371/20060822,AU6371DLF21
          'Combo1.AddItem "AU6371DFF25"  'AU6371DFT10/20070216
          'Combo1.AddItem "AU6371ELF25"  'AU6371EL/20061124
          'Combo1.AddItem "AU6371FLF25"  'AU6371S3/20061113
          'Combo1.AddItem "AU6371GLF25"   'AU6371GL
          'Combo1.AddItem "AU6371HLF25"   'AU6371HL
          'Combo1.AddItem "AU6371JLF25"   'AU6371EL


                Dim ChipString As String
                Dim i As Integer
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                
                
                
                ChipString = "vid_1984"
                
            
               If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
               
               ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
               '  CardResult = DO_WritePort(card, Channel_P1B, &H0)
                
                LBA = LBA + 1
                         
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                rv5 = 0
                rv6 = 0
                rv7 = 0
             
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                    If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 Call MsecDelay(0.05)
                 
                 
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1.2)   'power on time
              
                 '=========================================
                '    Noraml Mode Test
                '=========================================
                
                    If GetDeviceName(ChipString) = "" Then
                    Tester.Print "Normal mode Test Fail"
                    TestResult = "Bin2"
                    Call LabelMenu(0, 2, 1)
                                 
                    Exit Sub
                     End If
                
                
              
                '===============================================
                '  SD Card test
                '================================================
             '   If Left(ChipName, 10) = "AU6371DLF2" Then
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                
                  
                
                 
             '  End If
                   Call MsecDelay(0.01)
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(0.01)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      If rv0 = 1 Then
                          rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                          If rv0 <> 1 Then
                             rv0 = 2
                             Tester.Print "SD bus width Fail"
                          End If
                      End If
                          
                      
                      ClosePipe
                      
                      
                       For i = 1 To 20
                      
                        If rv0 = 1 Then
                           
                            ClosePipe
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                             
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                               
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                                                  
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                               
                            ClosePipe
                        End If
              
                        If rv0 <> 1 Then
                        GoTo AU6371DLResult
                        End If
                           
                        Next
                      
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                                    
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                     
                     
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                  If CardResult <> 0 Then
                    MsgBox "Set CF Card Detect On Fail"
                    End
                  End If
                  
                  
                    Call MsecDelay(0.01)
                 
                     CardResult = DO_WritePort(card, Channel_P1A, &H7D)
                 
               
                  
                
               
                 ClosePipe
                 rv1 = CBWTest_New(0, rv0, ChipString)
                 ClosePipe
  
                  
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
                   
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
         
                
                 
                 Call LabelMenu(1, rv1, rv0)
            
                      Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
              
                  
                  Call MsecDelay(0.01)
                 If rv1 = 1 Then
                     CardResult = DO_WritePort(card, Channel_P1A, &H7B)
                 End If
               
                 If CardResult <> 0 Then
                     MsgBox "Set SMC Card Detect Down Fail"
                     End
                  End If
                 
                 ClosePipe
                rv2 = CBWTest_New(0, rv1, "vid_058f")
                 Call LabelMenu(21, rv2, rv1)
                
                      Tester.Print rv2, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
              
              
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.01)
                If rv2 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                End If
                
                 
                   
                ClosePipe
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  MS Card test
                '================================================
                  
                     rv4 = rv3
               
                    ' Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect On Fail"
                    End
                 End If
                
                 Call MsecDelay(0.03)
                If rv4 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                End If
                   
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                 If rv5 = 1 Then
                          rv5 = Read_MS_Speed_AU6471(0, 0, 64, "4Bits")
                          If rv5 <> 1 Then
                             rv5 = 2
                             Tester.Print "MS bus width Fail"
                          End If
                      End If
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
               
                
               
                
                  CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
AU6371DLResult:
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
                        ElseIf rv2 = WRITE_FAIL Or rv3 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv2 = READ_FAIL Or rv3 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                         ElseIf rv4 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Or rv5 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                       
                            
                        ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub

Public Sub AU6471GLF23TestSub()
Tester.Print "AU6471GL is noraml mode"
     ' AU6371
         ' AF, CF : common board
         ' DL, EL, JL : common board
         ' GL,HL : common board
         ' EL,HL : card to open power

          'Combo1.AddItem "AU6371AFF25" ' AU6371S3/20061213
          'Combo1.AddItem "AU6371BLF25"   'AU6371
          'Combo1.AddItem "AU6371CFF25"  'AU6371CF/20061221
          'Combo1.AddItem "AU6371CLF25"  'AU6371EL
          'Combo1.AddItem "AU6371DLF25"  'AU6371/20060822,AU6371DLF21
          'Combo1.AddItem "AU6371DFF25"  'AU6371DFT10/20070216
          'Combo1.AddItem "AU6371ELF25"  'AU6371EL/20061124
          'Combo1.AddItem "AU6371FLF25"  'AU6371S3/20061113
          'Combo1.AddItem "AU6371GLF25"   'AU6371GL
          'Combo1.AddItem "AU6371HLF25"   'AU6371HL
          'Combo1.AddItem "AU6371JLF25"   'AU6371EL


                Dim ChipString As String
                Dim i As Integer
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                
                
                
                ChipString = "vid_058f"
                
            
               If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
               
               ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
               '  CardResult = DO_WritePort(card, Channel_P1B, &H0)
                
                LBA = LBA + 1
                         
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                rv5 = 0
                rv6 = 0
                rv7 = 0
             
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                    If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 Call MsecDelay(0.05)
                 
                 
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1.2)   'power on time
              
                 '=========================================
                '    Noraml Mode Test
                '=========================================
                
                    If GetDeviceName(ChipString) = "" Then
                    Tester.Print "Normal mode Test Fail"
                    TestResult = "Bin2"
                    Call LabelMenu(0, 2, 1)
                                 
                    Exit Sub
                     End If
                
                
              
                '===============================================
                '  SD Card test
                '================================================
             '   If Left(ChipName, 10) = "AU6371DLF2" Then
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                
                  
                
                 
             '  End If
                   Call MsecDelay(0.01)
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(0.01)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      If rv0 = 1 Then
                          rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                          If rv0 <> 1 Then
                             rv0 = 2
                             Tester.Print "SD bus width Fail"
                          End If
                      End If
                          
                      
                      ClosePipe
                      
                      
                       For i = 1 To 20
                      
                        If rv0 = 1 Then
                           
                            ClosePipe
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                             
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                               
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                                                  
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                               
                            ClosePipe
                        End If
              
                        If rv0 <> 1 Then
                        GoTo AU6371DLResult
                        End If
                           
                        Next
                      
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                                    
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                     
                     
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                  If CardResult <> 0 Then
                    MsgBox "Set CF Card Detect On Fail"
                    End
                  End If
                  
                  
                    Call MsecDelay(0.01)
                 
                     CardResult = DO_WritePort(card, Channel_P1A, &H7D)
                 
               
                  
                
               
                 ClosePipe
                 rv1 = CBWTest_New(0, rv0, ChipString)
                 ClosePipe
  
                  
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
                   
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
         
                
                 
                 Call LabelMenu(1, rv1, rv0)
            
                      Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
              
                  
                  Call MsecDelay(0.01)
                 If rv1 = 1 Then
                     CardResult = DO_WritePort(card, Channel_P1A, &H7B)
                 End If
               
                 If CardResult <> 0 Then
                     MsgBox "Set SMC Card Detect Down Fail"
                     End
                  End If
                 
                 ClosePipe
                rv2 = CBWTest_New(0, rv1, "vid_058f")
                 Call LabelMenu(21, rv2, rv1)
                
                      Tester.Print rv2, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
              
              
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.01)
                If rv2 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                End If
                
                 
                   
                ClosePipe
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  MS Card test
                '================================================
                  
                     rv4 = rv3
               
                    ' Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect On Fail"
                    End
                 End If
                
                 Call MsecDelay(0.03)
                If rv4 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                End If
                   
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                 If rv5 = 1 Then
                          rv5 = Read_MS_Speed_AU6471(0, 0, 64, "4Bits")
                          If rv5 <> 1 Then
                             rv5 = 2
                             Tester.Print "MS bus width Fail"
                          End If
                      End If
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
               
                
               
                
                  CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
AU6371DLResult:
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
                        ElseIf rv2 = WRITE_FAIL Or rv3 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv2 = READ_FAIL Or rv3 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                         ElseIf rv4 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Or rv5 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                       
                            
                        ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub

Public Sub AU6471GLF24TestSub()
Tester.Print "AU6471GL is noraml mode"
     ' AU6371
         ' AF, CF : common board
         ' DL, EL, JL : common board
         ' GL,HL : common board
         ' EL,HL : card to open power

          'Combo1.AddItem "AU6371AFF25" ' AU6371S3/20061213
          'Combo1.AddItem "AU6371BLF25"   'AU6371
          'Combo1.AddItem "AU6371CFF25"  'AU6371CF/20061221
          'Combo1.AddItem "AU6371CLF25"  'AU6371EL
          'Combo1.AddItem "AU6371DLF25"  'AU6371/20060822,AU6371DLF21
          'Combo1.AddItem "AU6371DFF25"  'AU6371DFT10/20070216
          'Combo1.AddItem "AU6371ELF25"  'AU6371EL/20061124
          'Combo1.AddItem "AU6371FLF25"  'AU6371S3/20061113
          'Combo1.AddItem "AU6371GLF25"   'AU6371GL
          'Combo1.AddItem "AU6371HLF25"   'AU6371HL
          'Combo1.AddItem "AU6371JLF25"   'AU6371EL


                Dim ChipString As String
                Dim i As Integer
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                
                
                
                ChipString = "pid_6366"
                
            
               If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
               
               ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
               '  CardResult = DO_WritePort(card, Channel_P1B, &H0)
                
                LBA = LBA + 1
                         
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                rv5 = 0
                rv6 = 0
                rv7 = 0
             
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &H80)
                    If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 Call MsecDelay(0.05)
                 
                 
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 'Call MsecDelay(1.2)   'power on time
                
                Call MsecDelay(0.3)
                rv0 = WaitDevOn(ChipString)
                Call MsecDelay(0.2)
                
                 '=========================================
                '    Noraml Mode Test
                '=========================================
                'If GetDeviceName(ChipString) = "" Then
                    
                If rv0 <> 1 Then
                    Tester.Print "Normal mode Test Fail"
                    TestResult = "Bin2"
                    Call LabelMenu(0, 2, 1)
                    Exit Sub
                End If
                
              
                '===============================================
                '  SD Card test
                '================================================
             '   If Left(ChipName, 10) = "AU6371DLF2" Then
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                
                  
                
                 
             '  End If
                   Call MsecDelay(0.01)
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(0.1)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      If rv0 = 1 Then
                          rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                          If rv0 <> 1 Then
                             rv0 = 2
                             Tester.Print "SD bus width Fail"
                          End If
                      End If
                          
                      
                      ClosePipe
                      
                      
                       For i = 1 To 20
                      
                        If rv0 = 1 Then
                           
                            ClosePipe
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                             
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                               
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                                                  
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                               
                            ClosePipe
                        End If
              
                        If rv0 <> 1 Then
                        GoTo AU6371DLResult
                        End If
                           
                        Next
                      
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                            UsbSpeedTestResult = GPO_FAIL
                            rv0 = 3
                          End If
                        End If
                     
                     
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                  If CardResult <> 0 Then
                    MsgBox "Set CF Card Detect On Fail"
                    End
                  End If
                  
                  
                    Call MsecDelay(0.01)
                 
                     CardResult = DO_WritePort(card, Channel_P1A, &H7D)
                 
               
                  
                
               
                 ClosePipe
                 rv1 = CBWTest_New(0, rv0, ChipString)
                 ClosePipe
  
                  
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
                   
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
         
                
                 
                 Call LabelMenu(1, rv1, rv0)
            
                      Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
              
                  
                  Call MsecDelay(0.01)
                 If rv1 = 1 Then
                     CardResult = DO_WritePort(card, Channel_P1A, &H7B)
                 End If
               
                 If CardResult <> 0 Then
                     MsgBox "Set SMC Card Detect Down Fail"
                     End
                  End If
                 
                 ClosePipe
                rv2 = CBWTest_New(0, rv1, "vid_058f")
                 Call LabelMenu(21, rv2, rv1)
                
                      Tester.Print rv2, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
              
              
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.01)
                If rv2 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                End If
                
                 
                   
                ClosePipe
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  MS Card test
                '================================================
                  
                     rv4 = rv3
               
                    ' Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect On Fail"
                    End
                 End If
                
                 Call MsecDelay(0.03)
                If rv4 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                End If
                   
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                 If rv5 = 1 Then
                          rv5 = Read_MS_Speed_AU6471(0, 0, 64, "4Bits")
                          If rv5 <> 1 Then
                             rv5 = 2
                             Tester.Print "MS bus width Fail"
                          End If
                      End If
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
               
AU6371DLResult:
                CardResult = DO_WritePort(card, Channel_P1A, &H80)   ' Close power
                
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
                        ElseIf rv2 = WRITE_FAIL Or rv3 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv2 = READ_FAIL Or rv3 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                         ElseIf rv4 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Or rv5 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                       
                            
                        ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub

Public Sub AU6471GLF04TestSub()

Dim HV_Done_Flag As Boolean
Dim HV_Result As String
Dim LV_Result As String

Tester.Print "AU6471GL is noraml mode"
     ' AU6371
         ' AF, CF : common board
         ' DL, EL, JL : common board
         ' GL,HL : common board
         ' EL,HL : card to open power

          'Combo1.AddItem "AU6371AFF25" ' AU6371S3/20061213
          'Combo1.AddItem "AU6371BLF25"   'AU6371
          'Combo1.AddItem "AU6371CFF25"  'AU6371CF/20061221
          'Combo1.AddItem "AU6371CLF25"  'AU6371EL
          'Combo1.AddItem "AU6371DLF25"  'AU6371/20060822,AU6371DLF21
          'Combo1.AddItem "AU6371DFF25"  'AU6371DFT10/20070216
          'Combo1.AddItem "AU6371ELF25"  'AU6371EL/20061124
          'Combo1.AddItem "AU6371FLF25"  'AU6371S3/20061113
          'Combo1.AddItem "AU6371GLF25"   'AU6371GL
          'Combo1.AddItem "AU6371HLF25"   'AU6371HL
          'Combo1.AddItem "AU6371JLF25"   'AU6371EL


                Dim ChipString As String
                Dim i As Integer
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                
                
                
                ChipString = "pid_6366"
                
            
'               If PCI7248InitFinish = 0 Then
'                  PCI7248Exist
'               End If
                If PCI7248InitFinish_Sync = 0 Then
                    PCI7248Exist_P1C_Sync
                End If
    

Routine_Label:
                
               ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
               '  CardResult = DO_WritePort(card, Channel_P1B, &H0)
                
                LBA = LBA + 1
                         
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                rv5 = 0
                rv6 = 0
                rv7 = 0
             
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
                
                
                '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &H80)
                    If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 Call MsecDelay(0.2)
                 
                 
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                
                If Not HV_Done_Flag Then
                    Call PowerSet2(0, "5.3", "0.5", 1, "5.3", "0.5", 1)
                    Call MsecDelay(0.3)
                    SetSiteStatus (RunHV)
                    Tester.Print "AU6471GL : HV Begin Test ..."
                Else
                    Call PowerSet2(0, "4.7", "0.5", 1, "4.7", "0.5", 1)
                    Call MsecDelay(0.4)
                    SetSiteStatus (RunLV)
                    Tester.Print vbCrLf & "AU6471GL : LV Begin Test ..."
                End If
                
                 'Call MsecDelay(1.2)   'power on time
                
                Call MsecDelay(0.3)
                rv0 = WaitDevOn(ChipString)
                Call MsecDelay(0.2)
                
                 '=========================================
                '    Noraml Mode Test
                '=========================================
                'If GetDeviceName(ChipString) = "" Then
                    
                If rv0 <> 1 Then
                    Tester.Print "Find Device Fail ..."
                    'TestResult = "Bin2"
                    Call LabelMenu(0, 2, 1)
                    GoTo AU6371DLResult
                End If
                
              
                '===============================================
                '  SD Card test
                '================================================
             '   If Left(ChipName, 10) = "AU6371DLF2" Then
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                
                  
                
                 
             '  End If
                   Call MsecDelay(0.01)
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(0.2)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      If rv0 = 1 Then
                          rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                          If rv0 <> 1 Then
                             rv0 = 2
                             Tester.Print "SD bus width Fail"
                          End If
                      End If
                          
                      
                      ClosePipe
                      
                      
                       For i = 1 To 20
                      
                        If rv0 = 1 Then
                           
                            ClosePipe
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                             
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                               
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                                                  
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                               
                            ClosePipe
                        End If
              
                        If rv0 <> 1 Then
                            ClosePipe
                            GoTo AU6371DLResult
                        End If
                           
                        Next
                      
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                            UsbSpeedTestResult = GPO_FAIL
                            rv0 = 3
                          End If
                        End If
                     
                     
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                  If CardResult <> 0 Then
                    MsgBox "Set CF Card Detect On Fail"
                    End
                  End If
                  
                  
                    Call MsecDelay(0.01)
                 
                     CardResult = DO_WritePort(card, Channel_P1A, &H7D)
                 
               
                  
                
               
                 ClosePipe
                 rv1 = CBWTest_New(0, rv0, ChipString)
                 ClosePipe
  
                  
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
                   
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
         
                
                 
                 Call LabelMenu(1, rv1, rv0)
            
                      Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
              
                  
                  Call MsecDelay(0.01)
                 If rv1 = 1 Then
                     CardResult = DO_WritePort(card, Channel_P1A, &H7B)
                 End If
               
                 If CardResult <> 0 Then
                     MsgBox "Set SMC Card Detect Down Fail"
                     End
                  End If
                 
                 ClosePipe
                rv2 = CBWTest_New(0, rv1, "vid_058f")
                 Call LabelMenu(21, rv2, rv1)
                
                      Tester.Print rv2, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
              
              
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.01)
                If rv2 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                End If
                
                 
                   
                ClosePipe
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  MS Card test
                '================================================
                  
                     rv4 = rv3
               
                    ' Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect On Fail"
                    End
                 End If
                
                 Call MsecDelay(0.03)
                If rv4 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                End If
                   
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                 If rv5 = 1 Then
                          rv5 = Read_MS_Speed_AU6471(0, 0, 64, "4Bits")
                          If rv5 <> 1 Then
                             rv5 = 2
                             Tester.Print "MS bus width Fail"
                          End If
                      End If
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
               
AU6371DLResult:
                If HV_Done_Flag = False Then
                    If rv0 * rv1 * rv2 * rv3 * rv4 * rv5 = 1 Then
                        SetSiteStatus (HVDone)
                    Else
                        SetSiteStatus (UNKNOW)
                    End If
                    Call WaitAnotherSiteDone(HVDone, 8#)
                Else
                    If rv0 * rv1 * rv2 * rv3 * rv4 * rv5 = 1 Then
                        SetSiteStatus (LVDone)
                    Else
                        SetSiteStatus (UNKNOW)
                    End If
                    Call WaitAnotherSiteDone(LVDone, 8#)
                End If
                
                CardResult = DO_WritePort(card, Channel_P1A, &H80)   ' Close power
                Call PowerSet2(0, "0.0", "0.5", 1, "0.0", "0.5", 1)
                WaitDevOFF (ChipString)
                Call MsecDelay(0.3)

                If HV_Done_Flag = False Then
                    If rv0 = 0 Then
                        HV_Result = "Bin2"
                        Tester.Print "HV Unknow"
                    ElseIf rv0 * rv1 * rv2 * rv3 * rv4 * rv5 <> 1 Then
                        HV_Result = "Fail"
                        Tester.Print "HV Fail"
                    ElseIf rv0 * rv1 * rv2 * rv3 * rv4 * rv5 = 1 Then
                        HV_Result = "PASS"
                        Tester.Print "HV PASS"
                    End If
                    
                    HV_Done_Flag = True
                    GoTo Routine_Label
                Else
                    If rv0 = 0 Then
                        LV_Result = "Bin2"
                        Tester.Print "LV Unknow"
                    ElseIf rv0 * rv1 * rv2 * rv3 * rv4 * rv5 <> 1 Then
                        LV_Result = "Fail"
                        Tester.Print "LV Fail"
                    ElseIf rv0 * rv1 * rv2 * rv3 * rv4 * rv5 = 1 Then
                        LV_Result = "PASS"
                        Tester.Print "LV PASS"
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
                
'                If rv0 = UNKNOW Then
'                     UnknowDeviceFail = UnknowDeviceFail + 1
'                     TestResult = "UNKNOW"
'                  ElseIf rv0 = WRITE_FAIL Then
'                      SDWriteFail = SDWriteFail + 1
'                      TestResult = "SD_WF"
'                  ElseIf rv0 = READ_FAIL Then
'                      SDReadFail = SDReadFail + 1
'                      TestResult = "SD_RF"
'                  ElseIf rv1 = WRITE_FAIL Then
'                      CFWriteFail = CFWriteFail + 1
'                      TestResult = "CF_WF"
'                  ElseIf rv1 = READ_FAIL Then
'                      CFReadFail = CFReadFail + 1
'                      TestResult = "CF_RF"
'                  ElseIf rv2 = WRITE_FAIL Or rv3 = WRITE_FAIL Then
'                      XDWriteFail = XDWriteFail + 1
'                      TestResult = "XD_WF"
'                  ElseIf rv2 = READ_FAIL Or rv3 = READ_FAIL Then
'                      XDReadFail = XDReadFail + 1
'                      TestResult = "XD_RF"
'                   ElseIf rv4 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
'                      MSWriteFail = MSWriteFail + 1
'                      TestResult = "MS_WF"
'                  ElseIf rv4 = READ_FAIL Or rv5 = READ_FAIL Then
'                      MSReadFail = MSReadFail + 1
'                      TestResult = "MS_RF"
'
'
'                  ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
'                       TestResult = "PASS"
'                  Else
'                      TestResult = "Bin2"
'
'                  End If
End Sub

Public Sub AU6471JLS10SortingSub()
Tester.Print "AU6471GL is noraml mode"
     ' AU6371
         ' AF, CF : common board
         ' DL, EL, JL : common board
         ' GL,HL : common board
         ' EL,HL : card to open power

          'Combo1.AddItem "AU6371AFF25" ' AU6371S3/20061213
          'Combo1.AddItem "AU6371BLF25"   'AU6371
          'Combo1.AddItem "AU6371CFF25"  'AU6371CF/20061221
          'Combo1.AddItem "AU6371CLF25"  'AU6371EL
          'Combo1.AddItem "AU6371DLF25"  'AU6371/20060822,AU6371DLF21
          'Combo1.AddItem "AU6371DFF25"  'AU6371DFT10/20070216
          'Combo1.AddItem "AU6371ELF25"  'AU6371EL/20061124
          'Combo1.AddItem "AU6371FLF25"  'AU6371S3/20061113
          'Combo1.AddItem "AU6371GLF25"   'AU6371GL
          'Combo1.AddItem "AU6371HLF25"   'AU6371HL
          'Combo1.AddItem "AU6371JLF25"   'AU6371EL


                Dim ChipString As String
                Dim i As Integer
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                
                
                
                ChipString = "vid_058f"
                
            
               If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
               
               ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
               '  CardResult = DO_WritePort(card, Channel_P1B, &H0)
                
                LBA = LBA + 1
                         
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                rv5 = 0
                rv6 = 0
                rv7 = 0
             
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                    If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 Call MsecDelay(0.05)
                 
                 
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1.2)   'power on time
              
                 '=========================================
                '    Noraml Mode Test
                '=========================================
                
                    If GetDeviceName(ChipString) = "" Then
                    Tester.Print "Normal mode Test Fail"
                    TestResult = "Bin2"
                    Call LabelMenu(0, 2, 1)
                                 
                    Exit Sub
                     End If
                
                
                 GoTo MSTest
                '===============================================
                '  SD Card test
                '================================================
             '   If Left(ChipName, 10) = "AU6371DLF2" Then
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                
                  
                
                 
             '  End If
                   Call MsecDelay(0.01)
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(0.01)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      If rv0 = 1 Then
                          rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                          If rv0 <> 1 Then
                             rv0 = 2
                             Tester.Print "SD bus width Fail"
                          End If
                      End If
                          
                      
                      ClosePipe
                      
                      
                       For i = 1 To 20
                      
                        If rv0 = 1 Then
                           
                            ClosePipe
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                             
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                               
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                                                  
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                               
                            ClosePipe
                        End If
              
                        If rv0 <> 1 Then
                        GoTo AU6371DLResult
                        End If
                           
                        Next
                      
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                                    
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                     
                     
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                  If CardResult <> 0 Then
                    MsgBox "Set CF Card Detect On Fail"
                    End
                  End If
                  
                  
                    Call MsecDelay(0.01)
                 
                     CardResult = DO_WritePort(card, Channel_P1A, &H7D)
                 
               
                  
                
               
                 ClosePipe
                 rv1 = CBWTest_New(0, rv0, ChipString)
                 ClosePipe
  
                  
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
                   
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
         
                
                 
                 Call LabelMenu(1, rv1, rv0)
            
                      Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
              
                  
                  Call MsecDelay(0.01)
                 If rv1 = 1 Then
                     CardResult = DO_WritePort(card, Channel_P1A, &H7B)
                 End If
               
                 If CardResult <> 0 Then
                     MsgBox "Set SMC Card Detect Down Fail"
                     End
                  End If
                 
                 ClosePipe
                rv2 = CBWTest_New(0, rv1, "vid_058f")
                 Call LabelMenu(21, rv2, rv1)
                
                      Tester.Print rv2, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
              
              
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.01)
                If rv2 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                End If
                
                 
                   
                ClosePipe
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  MS Card test
                '================================================
                  
                     rv4 = rv3
               
                    ' Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
MSTest:
                rv0 = 1
                rv1 = 1
                rv2 = 1
                rv3 = 1
                rv4 = 1
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect On Fail"
                    End
                 End If
                
                 Call MsecDelay(0.03)
                If rv4 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                End If
                   
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                 If rv5 = 1 Then
                          rv5 = Read_MS_Speed_AU6471(0, 0, 64, "1Bits")
                          If rv5 <> 1 Then
                             rv5 = 2
                             Tester.Print "MS bus width Fail"
                          End If
                      End If
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
               
                
               
                
                  CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
AU6371DLResult:
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
                        ElseIf rv2 = WRITE_FAIL Or rv3 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv2 = READ_FAIL Or rv3 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                         ElseIf rv4 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Or rv5 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                       
                            
                        ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub

Public Sub AU6471GLTest20()
     ' AU6371
         ' AF, CF : common board
         ' DL, EL, JL : common board
         ' GL,HL : common board
         ' EL,HL : card to open power

          'Combo1.AddItem "AU6371AFF25" ' AU6371S3/20061213
          'Combo1.AddItem "AU6371BLF25"   'AU6371
          'Combo1.AddItem "AU6371CFF25"  'AU6371CF/20061221
          'Combo1.AddItem "AU6371CLF25"  'AU6371EL
          'Combo1.AddItem "AU6371DLF25"  'AU6371/20060822,AU6371DLF21
          'Combo1.AddItem "AU6371DFF25"  'AU6371DFT10/20070216
          'Combo1.AddItem "AU6371ELF25"  'AU6371EL/20061124
          'Combo1.AddItem "AU6371FLF25"  'AU6371S3/20061113
          'Combo1.AddItem "AU6371GLF25"   'AU6371GL
          'Combo1.AddItem "AU6371HLF25"   'AU6371HL
          'Combo1.AddItem "AU6371JLF25"   'AU6371EL


                Dim ChipString As String
                Dim i As Integer
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
                ChipName = "AU6371DLF26"
                If Left(ChipName, 10) = "AU6371ELF2" Then
                  ChipName = "AU6371EL"
                  OldChipName = "AU6371ELF2"
                End If
                 
                 If Left(ChipName, 10) = "AU6371HLF2" Then
                  ChipName = "AU6371EL"
                  OldChipName = "AU6371HLF2"
                End If
                
                If ChipName = "AU6371S3" Then
                  ChipString = "vid_8751"
                Else
                  ChipString = "vid_058f"
                End If
                
                
                ChipString = "vid_1984"
                If ChipName = "AU6371EL" Or Left(ChipName, 10) = "AU6371GLF2" Then
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    If ChipName = "AU6371EL" Then
                        AU6371EL_BootTime = 1.2
                        If OldChipName = "AU6371ELF2" Then
                         AU6371EL_BootTime = 0.5
                        End If
                        
                    Else
                        AU6371EL_BootTime = 0.02
                    End If
                 Else
                    AU6371EL_SD = 0
                    AU6371EL_CF = 0
                    AU6371EL_XD = 0
                    AU6371EL_MS = 0
                    AU6371EL_MSP = 0
                    AU6371EL_BootTime = 0
                End If
            
               If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
               
               ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
               '  CardResult = DO_WritePort(card, Channel_P1B, &H0)
                
                LBA = LBA + 1
                         
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                rv5 = 0
                rv6 = 0
                rv7 = 0
             
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                    If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 Call MsecDelay(0.05)
                 
                 
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F - AU6371EL_SD)
                  
                 Call MsecDelay(1.2 + AU6371EL_BootTime)  'power on time
              
                '===============================================
                '  SD Card test
                '================================================
             '   If Left(ChipName, 10) = "AU6371DLF2" Then
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                
                  
                  
                 If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect On Fail"
                    End
                 End If
                 
             '  End If
                   Call MsecDelay(0.01)
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(0.01)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
                      
                      
                       For i = 1 To 20
                      
                        If rv0 = 1 Then
                           
                            ClosePipe
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                             
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                               
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                                                  
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                               
                            ClosePipe
                        End If
              
                        If rv0 <> 1 Then
                        GoTo AU6371DLResult
                        End If
                           
                        Next
                     If Left(ChipName, 10) = "AU6371DLF2" Then
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                                    
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                     ElseIf Left(ChipName, 7) = "AU6366C" Then
                          
                         If rv0 <> 0 Then
                          If LightOn <> 175 Or LightOff <> 255 Then
                                    
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                      
                     Else
                     
                          If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HBF Then
                                    
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                     End If
                     
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
                If ChipName <> "AU6371S3" Then
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                  If CardResult <> 0 Then
                    MsgBox "Set CF Card Detect On Fail"
                    End
                  End If
                  
                  
                    Call MsecDelay(0.01)
                 If rv0 = 1 Then
                     CardResult = DO_WritePort(card, Channel_P1A, &H7D)
                 End If
                 If CardResult <> 0 Then
                    MsgBox "Set CF Card Detect Down Fail"
                    End
                 End If
                  Call MsecDelay(AU6371EL_BootTime * 2)
                 If ChipName = "AU6371EL" Then
                   ReaderExist = 0
                 End If
               
                 ClosePipe
                 rv1 = CBWTest_New(0, rv0, ChipString)
                 ClosePipe
  
                  
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
                   
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, ChipString)
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                
                
                
                 
               Else
                  rv1 = 1  '----------- AU6371S3 dp not have CF slot
                 
               End If
                 
                 Call LabelMenu(1, rv1, rv0)
            
                      Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              '    CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
              '   If CardResult <> 0 Then
              '      MsgBox "Set SMC Card Detect On Fail"
              '      End
              '   End If
                  
              '   Call MsecDelay(0.01)
              '  If rv1 = 1 Then
              '      CardResult = DO_WritePort(card, Channel_P1A, &H7B)
              '  End If
               
              '  If CardResult <> 0 Then
              '      MsgBox "Set SMC Card Detect Down Fail"
              '      End
              '   End If
                 
              '  ClosePipe
              '  rv2 = CBWTest_New(0, rv1, "vid_058f")
              '  Call LabelMenu(21, rv2, rv1)
                
              '       Tester.print rv2, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
              
               If rv1 = 1 Then
                   rv2 = 1   ' to complete the SMC asbolish
               Else
                   rv2 = 0
               End If
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.01)
                If rv2 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                End If
                Call MsecDelay(AU6371EL_BootTime)
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect Down Fail"
                    End
                 End If
                 
                   If ChipName = "AU6371EL" Then
                   ReaderExist = 0
                 End If
                
                ClosePipe
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  MS Card test
                '================================================
                  
                     rv4 = 1
               
               
                     Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect On Fail"
                    End
                 End If
                
                 Call MsecDelay(0.03)
                If rv4 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                End If
                  Call MsecDelay(AU6371EL_BootTime * 2)
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect Down Fail"
                    End
                 End If
                 If ChipName = "AU6371EL" Then
                   ReaderExist = 0
                 End If
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
               
                
               
                
                  CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
AU6371DLResult:
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
                        ElseIf rv2 = WRITE_FAIL Or rv3 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv2 = READ_FAIL Or rv3 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                         ElseIf rv4 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Or rv5 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                       
                            
                        ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub
