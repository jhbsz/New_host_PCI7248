Attribute VB_Name = "AU6430MDL"

Public Sub AU6430ELF22NormalTestSub()

Tester.Print "AU6430EL is NB mode"
Tester.Print "use v1.3 socket borad"

                Dim ChipString As String
                Dim i As Integer
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
                ChipString = "vid_058f"
           
                AU6371EL_BootTime = 0.3
              
       '1. power on intital
       
               If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
               
         
                
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
                 
                 
                 
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on
                  
                 Call MsecDelay(1.2 + AU6371EL_BootTime)  'power on time
                 
                '==============================================
                '  print NB mode test
                '==============================================
                
                If GetDeviceName(ChipString) <> "" Then
                    Tester.Print "NB mode Test Fail"
                    TestResult = "Bin2"
                    Call LabelMenu(0, 2, 1)
                                 
                    Exit Sub
                End If
                
            
              
                '===============================================
                '  Test light off
                '================================================
                     Call MsecDelay(0.01)
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                 
                 If LightOff <> 255 Then
                    Tester.Print Hex(LightOff); "   Light OFF Test Fail"
                    TestResult = "Bin3"
                  Exit Sub
                End If
              
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                    CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1.2)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                      If LightOn <> 191 Then
                         Tester.Print Hex(LightOn); "   Light ON Test Fail"
                         TestResult = "Bin3"
                         Exit Sub
                      End If
                     
               
                      ClosePipe
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
           
                    
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                     Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                  '===============================================
                   '  SMC Card test
                   '================================================
                    CardResult = DO_WritePort(card, Channel_P1A, &H7F)  'SMC +SD
                    Call MsecDelay(0.1)
                    
                    CardResult = DO_WritePort(card, Channel_P1A, &H7B)  'SMC
                    Call MsecDelay(2.1)
    
                   
                    If CardResult <> 0 Then
                        MsgBox "Set SMC Card Detect Down Fail"
                     End
                    End If
    
                   ClosePipe
                   rv1 = CBWTest_New(0, 1, ChipString)
                   ClosePipe
                  Call LabelMenu(1, rv1, rv0)
                  
                     Tester.Print rv1, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
               
                  '===============================================
                  '  XD Card test
                  '================================================
                  
                   CardResult = DO_WritePort(card, Channel_P1A, &H7F)  'XD + SMC
                   Call MsecDelay(0.1)
   
                  
                   
                   CardResult = DO_WritePort(card, Channel_P1A, &H77)  'XD
                   Call MsecDelay(2.1)
   
                  
                   If CardResult <> 0 Then
                       MsgBox "Set XD Card Detect Down Fail"
                    End
                   End If
   
                 ClosePipe
                rv2 = CBWTest_New(0, 1, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv2, rv1)
                 
                     Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
            
    
                 '===============================================
                '  MS Pro Card test
                '================================================
              
                    CardResult = DO_WritePort(card, Channel_P1A, &H7F)  'MS + XD
                   Call MsecDelay(0.1)
   
                  
                   CardResult = DO_WritePort(card, Channel_P1A, &H5F)  'MS
                   Call MsecDelay(2.1)
   
                  
                   If CardResult <> 0 Then
                       MsgBox "Set XD Card Detect Down Fail"
                    End
                   End If
   
                 ClosePipe
                rv3 = CBWTest_New(0, 1, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                
                Tester.Print rv3, " \\MSPRO :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
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

Public Sub AU6430BLF22TestSub()

Tester.Print "AU6430BL is Normal  mode; CIS Disable"
Tester.Print "use v1.3 socket borad"

                Dim ChipString As String
                Dim i As Integer
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
                ChipString = "vid_058f"
           
                AU6371EL_BootTime = 0.3
              
       '1. power on intital
       
               If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
               
         
                
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
                 
                 
                 
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on
                  
                 Call MsecDelay(1.2 + AU6371EL_BootTime)  'power on time
                 
                '==============================================
                '  print NB mode test
                '==============================================
                
                If GetDeviceName(ChipString) = "" Then
                    Tester.Print "Normal mode Test Fail"
                    TestResult = "Bin2"
                    Call LabelMenu(0, 2, 1)
                                 
                    Exit Sub
                End If
                
            
              
                '===============================================
                '  Test light off
                '================================================
                     Call MsecDelay(0.01)
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                 
                 If LightOff <> 255 Then
                    Tester.Print Hex(LightOff); "   Light OFF Test Fail"
                    TestResult = "Bin3"
                  Exit Sub
                End If
              
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                    CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1.2)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                      If LightOn <> 191 Then
                         Tester.Print Hex(LightOn); "   Light ON Test Fail"
                         TestResult = "Bin3"
                         Exit Sub
                      End If
                     
               
                      ClosePipe
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
           
                    
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                     Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                  '===============================================
                   '  SMC Card test
                   '================================================
                    CardResult = DO_WritePort(card, Channel_P1A, &H7F)  'SMC +SD
                    Call MsecDelay(0.2)
                    
                    CardResult = DO_WritePort(card, Channel_P1A, &H7B)  'SMC
                    Call MsecDelay(0.2)
    
                   
                    If CardResult <> 0 Then
                        MsgBox "Set SMC Card Detect Down Fail"
                     End
                    End If
    
                   ClosePipe
                   rv1 = CBWTest_New(0, rv0, ChipString)
                   ClosePipe
                  Call LabelMenu(1, rv1, rv0)
                   If rv1 <> 1 Then
                        Tester.Label9.Caption = "SMC Fail"
                     End If
                     Tester.Print rv1, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
               
                  '===============================================
                  '  XD Card test
                  '================================================
                  
                   CardResult = DO_WritePort(card, Channel_P1A, &H7F)  'XD + SMC
                   Call MsecDelay(0.2)
   
                  
                   
                   CardResult = DO_WritePort(card, Channel_P1A, &H77)  'XD
                   Call MsecDelay(0.2)
   
                  
                   If CardResult <> 0 Then
                       MsgBox "Set XD Card Detect Down Fail"
                    End
                   End If
   
                 ClosePipe
                rv2 = CBWTest_New(0, rv1, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv2, rv1)
                 
                     Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
            
    
                 '===============================================
                '  MS Pro Card test
                '================================================
              
                    CardResult = DO_WritePort(card, Channel_P1A, &H7F)  'MS + XD
                   Call MsecDelay(0.2)
   
                  
                   CardResult = DO_WritePort(card, Channel_P1A, &H5F)  'MS
                   Call MsecDelay(0.2)
   
                  
                   If CardResult <> 0 Then
                       MsgBox "Set MS Card Detect Down Fail"
                    End
                   End If
   
                 ClosePipe
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(3, rv3, rv2)
                
                Tester.Print rv3, " \\MSPRO :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
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
                       
                            
                        ElseIf rv2 = WRITE_FAIL Or rv1 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv2 = READ_FAIL Or rv1 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                         ElseIf rv3 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv3 = READ_FAIL Or rv5 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                       
                            
                        ElseIf rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub
Public Sub AU6430QLF22TestSub()

Tester.Print "AU6430QL is Normal  mode; CIS enable"
Tester.Print "use v1.3 socket borad"

                Dim ChipString As String
                Dim i As Integer
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
                ChipString = "vid_058f"
           
                AU6371EL_BootTime = 0.3
              
       '1. power on intital
       
               If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
               
         
                
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
                 
                 
                 
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on
                  
                 Call MsecDelay(1.2 + AU6371EL_BootTime)  'power on time
                 
                '==============================================
                '  print NB mode test
                '==============================================
                
                If GetDeviceName(ChipString) = "" Then
                    Tester.Print "Normal mode Test Fail"
                    TestResult = "Bin2"
                    Call LabelMenu(0, 2, 1)
                                 
                    Exit Sub
                End If
                
            
              
                '===============================================
                '  Test light off
                '================================================
                     Call MsecDelay(0.01)
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                 
                 If LightOff <> 255 Then
                    Tester.Print Hex(LightOff); "   Light OFF Test Fail"
                    TestResult = "Bin3"
                  Exit Sub
                End If
              
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                    CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(0.3)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                      If LightOn <> 191 Then
                         Tester.Print Hex(LightOn); "   Light ON Test Fail"
                         TestResult = "Bin3"
                         Exit Sub
                      End If
                     
               
                      ClosePipe
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
           
                    
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                     Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                  '===============================================
                   '  SMC Card test
                   '================================================
                    CardResult = DO_WritePort(card, Channel_P1A, &H7F)  'SMC +SD
                    Call MsecDelay(0.2)
                    
                    CardResult = DO_WritePort(card, Channel_P1A, &H7B)  'SMC
                    Call MsecDelay(0.2)
    
                   
                    If CardResult <> 0 Then
                        MsgBox "Set SMC Card Detect Down Fail"
                     End
                    End If
    
                   ClosePipe
                   rv1 = CBWTest_New_CIS(0, rv0, ChipString)
                   ClosePipe
                  Call LabelMenu(1, rv1, rv0)
                   If rv1 <> 1 Then
                        Tester.Label9.Caption = "SMC Fail"
                     End If
                     Tester.Print rv1, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
               
                  '===============================================
                  '  XD Card test
                  '================================================
                  
                   CardResult = DO_WritePort(card, Channel_P1A, &H7F)  'XD + SMC
                   Call MsecDelay(0.2)
   
                  
                   
                   CardResult = DO_WritePort(card, Channel_P1A, &H77)  'XD
                   Call MsecDelay(0.2)
   
                  
                   If CardResult <> 0 Then
                       MsgBox "Set XD Card Detect Down Fail"
                    End
                   End If
   
                 ClosePipe
                rv2 = CBWTest_New(0, rv1, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv2, rv1)
                 
                     Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
            
    
                 '===============================================
                '  MS Pro Card test
                '================================================
              
                    CardResult = DO_WritePort(card, Channel_P1A, &H7F)  'MS + XD
                   Call MsecDelay(0.2)
   
                  
                   CardResult = DO_WritePort(card, Channel_P1A, &H5F)  'MS
                   Call MsecDelay(0.2)
   
                  
                   If CardResult <> 0 Then
                       MsgBox "Set MS Card Detect Down Fail"
                    End
                   End If
   
                 ClosePipe
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(3, rv3, rv2)
                
                Tester.Print rv3, " \\MSPRO :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
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
                       
                            
                        ElseIf rv2 = WRITE_FAIL Or rv1 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv2 = READ_FAIL Or rv1 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                         ElseIf rv3 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv3 = READ_FAIL Or rv5 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                       
                            
                        ElseIf rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub
Public Sub AU6430ELF22TestSub()

Tester.Print "AU6430EL is NB mode ; CIS Disable"
Tester.Print "use v1.3 socket borad"

                Dim ChipString As String
                Dim i As Integer
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
                ChipString = "vid_058f"
           
                AU6371EL_BootTime = 0.3
              
       '1. power on intital
       
               If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
               
         
                
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
                 
                 
                 
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on
                  
                 Call MsecDelay(1.2 + AU6371EL_BootTime)  'power on time
                 
                '==============================================
                '  print NB mode test
                '==============================================
                
                If GetDeviceName(ChipString) <> "" Then
                    Tester.Print "NB mode Test Fail"
                    TestResult = "Bin2"
                    Call LabelMenu(0, 2, 1)
                                 
                    Exit Sub
                End If
                
            
              
                '===============================================
                '  Test light off
                '================================================
                     Call MsecDelay(0.01)
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                 
                 If LightOff <> 255 Then
                    Tester.Print Hex(LightOff); "   Light OFF Test Fail"
                    TestResult = "Bin3"
                  Exit Sub
                End If
              
                 '===========================================
                 'SD Test
                 '============================================
  
                     ' set SD card detect down
                    CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1.2)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                      If LightOn <> 191 Then
                         Tester.Print Hex(LightOn); "   Light ON Test Fail"
                         TestResult = "Bin3"
                         Exit Sub
                      End If
                     
               
                      ClosePipe
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
           
                    
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                     Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                  '===============================================
                   '  SMC Card test  : CIS Diable , rv1 =1 is pass
                   '================================================
                    CardResult = DO_WritePort(card, Channel_P1A, &H7A)  'SMC +SD
                    Call MsecDelay(0.2)
                    
                    CardResult = DO_WritePort(card, Channel_P1A, &H7B)  'SMC
                    Call MsecDelay(0.2)
                    OpenPipe
                    rv1 = ReInitial(0)
                    ClosePipe
                    If CardResult <> 0 Then
                        MsgBox "Set SMC Card Detect Down Fail"
                     End
                    End If
    
                   ClosePipe
                   rv1 = CBWTest_New(0, rv0, ChipString) ' can not use CIS check, otherwise chip will fail at XD R/W
                   ClosePipe
                  Call LabelMenu(1, rv1, rv0)
                  
                     Tester.Print rv1, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                     If rv1 <> 1 Then
                        Tester.Label9.Caption = "SMC Fail"
                     End If
                  '===============================================
                  '  XD Card test
                  '================================================
                  
                   CardResult = DO_WritePort(card, Channel_P1A, &H73)  'XD + SMC
                   Call MsecDelay(0.2)
   
                  
                   
                   CardResult = DO_WritePort(card, Channel_P1A, &H77)  'XD
                   Call MsecDelay(0.2)
   
                     OpenPipe
                    rv2 = ReInitial(0)
                    ClosePipe
                   If CardResult <> 0 Then
                       MsgBox "Set XD Card Detect Down Fail"
                    End
                   End If
   
                 ClosePipe
                rv2 = CBWTest_New(0, rv1, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv2, rv1)
                 
                     Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
            
    
                 '===============================================
                '  MS Pro Card test
                '================================================
              
                    CardResult = DO_WritePort(card, Channel_P1A, &H57)  'MS + XD
                   Call MsecDelay(0.2)
   
                  
                   CardResult = DO_WritePort(card, Channel_P1A, &H5F)  'MS
                   Call MsecDelay(0.2)
   
                     OpenPipe
                    rv3 = ReInitial(0)
                    ClosePipe
                   If CardResult <> 0 Then
                       MsgBox "Set MS Card Detect Down Fail"
                    End
                   End If
   
                 ClosePipe
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                
                Tester.Print rv3, " \\MSPRO :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
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
                       
                            
                        ElseIf rv2 = WRITE_FAIL Or rv1 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv2 = READ_FAIL Or rv1 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                         ElseIf rv3 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv3 = READ_FAIL Or rv5 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                       
                            
                        ElseIf rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub

Public Sub AU6430DLF22TestSub()

Tester.Print "AU6430DL is NB mode ; CIS enable"
Tester.Print "use v1.3 socket borad"

                Dim ChipString As String
                Dim i As Integer
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
                ChipString = "vid_058f"
           
                AU6371EL_BootTime = 0.3
              
       '1. power on intital
       
               If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
               
         
                
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
                 
                 
                 
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on
                  
                 Call MsecDelay(1.2 + AU6371EL_BootTime)  'power on time
                 
                '==============================================
                '  print NB mode test
                '==============================================
                
                If GetDeviceName(ChipString) <> "" Then
                    Tester.Print "NB mode Test Fail"
                    TestResult = "Bin2"
                    Call LabelMenu(0, 2, 1)
                                 
                    Exit Sub
                End If
                
            
              
                '===============================================
                '  Test light off
                '================================================
                     Call MsecDelay(0.01)
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                 
                 If LightOff <> 255 Then
                    Tester.Print Hex(LightOff); "   Light OFF Test Fail"
                    TestResult = "Bin3"
                  Exit Sub
                End If
              
                 '===========================================
                 'SD Test
                 '============================================
  
                     ' set SD card detect down
                    CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1.2)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                      If LightOn <> 191 Then
                         Tester.Print Hex(LightOn); "   Light ON Test Fail"
                         TestResult = "Bin3"
                         Exit Sub
                      End If
                     
               
                      ClosePipe
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
           
                    
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                     Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                  '===============================================
                   '  SMC Card test  : CIS enable , rv1 =1 is pass
                   '================================================
                    CardResult = DO_WritePort(card, Channel_P1A, &H7A)  'SMC +SD
                    Call MsecDelay(0.2)
                    
                    CardResult = DO_WritePort(card, Channel_P1A, &H7B)  'SMC
                    Call MsecDelay(0.2)
                    OpenPipe
                    rv1 = ReInitial(0)
                    ClosePipe
                    If CardResult <> 0 Then
                        MsgBox "Set SMC Card Detect Down Fail"
                     End
                    End If
    
                   ClosePipe
                   rv1 = CBWTest_New_CIS(0, rv0, ChipString)
                   ClosePipe
                  Call LabelMenu(1, rv1, rv0)
                  
                     Tester.Print rv1, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                     If rv1 <> 1 Then
                        Tester.Label9.Caption = "SMC Fail"
                     End If
                  '===============================================
                  '  XD Card test
                  '================================================
                  
                   CardResult = DO_WritePort(card, Channel_P1A, &H73)  'XD + SMC
                   Call MsecDelay(0.2)
   
                  
                   
                   CardResult = DO_WritePort(card, Channel_P1A, &H77)  'XD
                   Call MsecDelay(0.2)
   
                     OpenPipe
                    rv2 = ReInitial(0)
                    ClosePipe
                   If CardResult <> 0 Then
                       MsgBox "Set XD Card Detect Down Fail"
                    End
                   End If
   
                 ClosePipe
                rv2 = CBWTest_New(0, rv1, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv2, rv1)
                 
                     Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
            
    
                 '===============================================
                '  MS Pro Card test
                '================================================
              
                    CardResult = DO_WritePort(card, Channel_P1A, &H57)  'MS + XD
                   Call MsecDelay(0.2)
   
                  
                   CardResult = DO_WritePort(card, Channel_P1A, &H5F)  'MS
                   Call MsecDelay(0.2)
   
                     OpenPipe
                    rv3 = ReInitial(0)
                    ClosePipe
                   If CardResult <> 0 Then
                       MsgBox "Set MS Card Detect Down Fail"
                    End
                   End If
   
                 ClosePipe
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                
                Tester.Print rv3, " \\MSPRO :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
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
                       
                            
                        ElseIf rv2 = WRITE_FAIL Or rv1 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv2 = READ_FAIL Or rv1 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                         ElseIf rv3 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv3 = READ_FAIL Or rv5 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                       
                            
                        ElseIf rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub
Public Sub AU6430QLF21TestSub()

Tester.Print "AU6430QL is NB mode"

                Dim ChipString As String
                Dim i As Integer
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
                ChipString = "vid_058f"
           
                AU6371EL_BootTime = 0.3
              
            
               If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
               
         
                
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
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on
                  
                 Call MsecDelay(1.2 + AU6371EL_BootTime)  'power on time
                 
                '==============================================
                '  print NB mode test
                '==============================================
                
                
                
                 
                 
              
                '===============================================
                '  SD Card test
                '================================================
              
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                
                  
                  
                 If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect On Fail"
                    End
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
                     Call MsecDelay(0.01)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
               
                      ClosePipe
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
                      
                      
                      
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
                
                 If rv0 = 1 Then
                  rv1 = 1  '----------- AU6371S3 dp not have CF slot
                 Else
                   rv1 = 4
                 End If
                 
               '  Call LabelMenu(1, rv1, rv0)
            
                 '     Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
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
                   rv2 = 4
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
                  
                  If rv3 = 1 Then
                     rv4 = 1
                  Else
                    rv4 = 4
                  End If
                      
               
                   '  Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
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



Public Sub AU6430MPSub()

          

             If ChipName = "AU6430QLF21" Then
                 Call AU6430QLF21TestSub
              

            ElseIf ChipName = "AU6430QLF20" Then
                 Call AU6430QLTestSub
             
              
            ElseIf ChipName = "AU6430QLS10" Or ChipName = "AU6430ELS10" Or ChipName = "AU6430DLS10" Then
                 Call AU6430QLS10SortingSub
              
              
              
             ElseIf ChipName = "AU6430QLS11" Or ChipName = "AU6430ELS11" Or ChipName = "AU6430DLS11" Then
                 Call AU6430QLS11SortingSub
               
                 
             ElseIf ChipName = "AU6430QLS12" Or ChipName = "AU6430ELS12" Or ChipName = "AU6430DLS12" Then
                 Call AU6430QLS12SortingSub
              
                 
              ElseIf ChipName = "AU698XHLS10" Or ChipName = "AU6430QLS13" Or ChipName = "AU6430ELS13" Or ChipName = "AU6430DLS13" Then
                 Call AU6430QLS13SortingSub
             
              
              ElseIf ChipName = "AU698XHLS20" Then
                    Call AU698XHLS20SortingSub
              
                
                
                ElseIf ChipName = "AU6337BLS10" Then
                    Call AU6337BLS10SortingSub
              
              
               ElseIf ChipName = "AU6430DLF20" Then
            
                 Call AU6430DLTestSub
               
               
               ElseIf ChipName = "AU6430ELF20" Then
            
                 Call AU6430ELTestSub
               
                 
             
                 
                 
                ElseIf ChipName = "AU6430BLF20" Then
            
                 Call AU6430BLTestSub
               
               
                  ElseIf ChipName = "AU6430ELS20" Then
                 Call AU6430ELSortingSub
                 
                 
                 
             '===============================================
             
                ElseIf ChipName = "AU6430QLF22" Then
                 Call AU6430QLF22TestSub
              
              
                ElseIf ChipName = "AU6430ELF22" Then
            
                 Call AU6430ELF22TestSub
               
               
                ElseIf ChipName = "AU6430BLF22" Then
            
                 Call AU6430BLF22TestSub
               
               
                  ElseIf ChipName = "AU6430DLF22" Then
            
                  Call AU6430DLF22TestSub
               
               
                ElseIf ChipName = "AU6433DFS10" Then
            
                  Call AU6433DFS10SortingSub
               
               
                   ElseIf ChipName = "AU6471FLS10" Then
            
                   Call AU6471FLS10SortingSub
                
                 
                    ElseIf ChipName = "AU6471FLS11" Then
            
                   Call AU6471FLS11SortingSub
                
                 
               ElseIf ChipName = "AU6433EFS10" Then
            
                  Call AU6433EFS10SortingSub
               
               
               
                ElseIf ChipName = "AU6476BLS10" Then
            
                   Call AU6433EFS10SortingSub
                
                
                
                 ElseIf ChipName = "AU6430BLS10" Then
            
                   Call AU6433EFS10SortingSub
                
                
                 ElseIf ChipName = "AU6256XLS10" Then
            
                   Call AU6256XLS10SortingSub
                
                
                  ElseIf ChipName = "AU6256XLS11" Then
            
                   Call AU6256XLS11SortingSub
                
                
                  ElseIf ChipName = "AU6256XLS12" Then
            
                   Call AU6256XLS12SortingSub
                
                
                  ElseIf ChipName = "AU6256XLS13" Then
            
                   Call AU6256XLS13SortingSub
                
                
                ElseIf ChipName = "AU6256XLS14" Then
            
                   Call AU6256XLS14SortingSub
                   
                 ElseIf ChipName = "AU6256XLS19" Then
            
                   Call AU6256XLS19SortingSub
                   
                   ElseIf ChipName = "AU6256XLS1A" Then
            
                   Call AU6256XLS1ASortingSub
                   
                    ElseIf ChipName = "AU6256XLS1B" Then
            
                   Call AU6256XLS1BSortingSub
                   
                   ElseIf ChipName = "AU6256XLS1C" Then
            
                   Call AU6256XLS1CSortingSub
                   
                    ElseIf ChipName = "AU6256XLS1D" Then
            
                   Call AU6256XLS1DSortingSub
                   
                   ElseIf ChipName = "AU6256XLS20" Then
            
                   Call AU6256XLS20SortingSub
                   
                   ElseIf ChipName = "AU6256XLS21" Then
            
                   Call AU6256XLS21SortingSub
                   
                   ElseIf ChipName = "AU6256XLS31" Then
            
                   Call AU6256XLS31SortingSub
                   
                   ElseIf ChipName = "AU6256XLS32" Then
            
                   Call AU6256XLS32SortingSub
                   
                   ElseIf ChipName = "AU6256XLSE2" Then
            
                   Call AU6256XLSE2SortingSub
                   
                   ElseIf ChipName = "AU6256XLS15" Then
            
                   Call AU6256XLS15SortingSub
                   
                   
                   ElseIf ChipName = "AU6256XLS16" Then
            
                   Call AU6256XLS16SortingSub
                   
                   ElseIf ChipName = "AU6256XLS17" Then
            
                   Call AU6256XLS17SortingSub
                   
                   ElseIf ChipName = "AU6256XLS18" Then
            
               '    Call AU6256XLS18SortingSub
                   
                End If
End Sub

Public Sub AU6430ELTestSub()
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
                
                If Left(ChipName, 10) = "AU6430QLF2" Then
                   ChipName = "AU6371DLF20"
                 
                End If
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
             
                  CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                 Call MsecDelay(1.5 + AU6371EL_BootTime)  'power on time
              
                '===============================================
                '  SD Card test
                '================================================
             
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
                      
                      
                      
                     If Left(ChipName, 10) = "AU6371DLF2" Then
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Then
                                    
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
                
                 If rv0 = 1 Then
                  rv1 = 1  '----------- AU6371S3 dp not have CF slot
                 Else
                   rv1 = 4
                 End If
                 
               '  Call LabelMenu(1, rv1, rv0)
            
                 '     Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
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
                   rv2 = 4
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
                Call MsecDelay(AU6371EL_BootTime + 1.2)
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect Down Fail"
                    End
                 End If
                 
                  
                   ReaderExist = 0
                 
                ClosePipe
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  MS Card test
                '================================================
                  
                  If rv3 = 1 Then
                     rv4 = 1
                  Else
                    rv4 = 4
                  End If
                      
               
                   '  Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
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
                  Call MsecDelay(AU6371EL_BootTime + 1)
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect Down Fail"
                    End
                 End If
               
                ReaderExist = 0
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

Public Sub AU6430ELSortingSub()
 
' program for MS card unstable
 
 
ChipName = "AU6371ELF20"
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
                Dim TmpLBA As Long
                OldChipName = ""
                
                If Left(ChipName, 10) = "AU6430QLF2" Then
                   ChipName = "AU6371DLF20"
                 
                End If
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
               '  Call MsecDelay(0.05)
                 
                 
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
             
                  CardResult = DO_WritePort(card, Channel_P1A, &H7E)
              '   Call MsecDelay(1.5 + AU6371EL_BootTime)  'power on time
              
                '===============================================
                '  SD Card test
                '================================================
             
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
                      
                      
                      
                     If Left(ChipName, 10) = "AU6371DLF2" Then
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Then
                                    
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
                     
                        '  If rv0 <> 0 Then
                        '  If LightON <> &HBF Or LightOFF <> &HBF Then
                                    
                        '  UsbSpeedTestResult = GPO_FAIL
                        '  rv0 = 3
                        '  End If
                        'End If
                     End If
                     
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
                
                 If rv0 = 1 Then
                  rv1 = 1  '----------- AU6371S3 dp not have CF slot
                 Else
                   rv1 = 4
                 End If
                 
               '  Call LabelMenu(1, rv1, rv0)
            
                 '     Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
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
                   rv2 = 4
               End If
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
              '   Call MsecDelay(0.01)
                If rv2 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                End If
             '   Call MsecDelay(AU6371EL_BootTime + 1.2)
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect Down Fail"
                    End
                 End If
                 
                  
                   ReaderExist = 0
                 
                ClosePipe
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  MS Card test
                '================================================
                  
                  If rv3 = 1 Then
                     rv4 = 1
                  Else
                    rv4 = 4
                  End If
                      
               
                   '  Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect On Fail"
                    End
                 End If
                
              '   Call MsecDelay(0.03)
                If rv4 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                End If
                  Call MsecDelay(AU6371EL_BootTime + 1)
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect Down Fail"
                    End
                 End If
               
                ReaderExist = 0
                ClosePipe
                i = 0
                Do
                rv5 = 0
                rv5 = CBWTest_New(0, rv4, ChipString)
                i = i + 1
                Loop While rv5 = 1 And i < 50
                 
                Tester.Print rv5, " \\MSpro short pattern :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                 Tester.Print i
                '========== long pattern
                If rv5 = 1 Then
                TmpLBA = LBA
                 For i = 1 To 150
                             rv5 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv5 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv5 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6371DLResult
                             End If
                  Next
                
                
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro long pattern :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                      Tester.Print i
                ClosePipe
               
                
               End If
                
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
 

Public Sub AU6430BLTestSub()
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
                
                 If Left(ChipName, 10) = "AU6430QLF2" Then
                  ChipName = "AU6371DLF20"
                 
                End If
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
             
                  CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                 Call MsecDelay(1.5 + AU6371EL_BootTime)  'power on time
              
                '===============================================
                '  SD Card test
                '================================================
             
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
                      
                      
                      
                     If Left(ChipName, 10) = "AU6371DLF2" Then
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Then
                                    
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
                
                 If rv0 = 1 Then
                  rv1 = 1  '----------- AU6371S3 dp not have CF slot
                 Else
                   rv1 = 4
                 End If
                 
               '  Call LabelMenu(1, rv1, rv0)
            
                 '     Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
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
                   rv2 = 4
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
                Call MsecDelay(AU6371EL_BootTime + 1.2)
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect Down Fail"
                    End
                 End If
                 
                  
                   ReaderExist = 0
                 
                ClosePipe
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  MS Card test
                '================================================
                  
                  If rv3 = 1 Then
                     rv4 = 1
                  Else
                    rv4 = 4
                  End If
                      
               
                   '  Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
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
                  Call MsecDelay(AU6371EL_BootTime + 1)
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect Down Fail"
                    End
                 End If
               
                ReaderExist = 0
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

Public Sub AU6430DLTestSub()
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
                
                 If Left(ChipName, 10) = "AU6430QLF2" Then
                  ChipName = "AU6371DLF20"
                 
                End If
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
             
                  CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                 Call MsecDelay(1.5 + AU6371EL_BootTime)  'power on time
              
                '===============================================
                '  SD Card test
                '================================================
             
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
                      
                      
                      
                     If Left(ChipName, 10) = "AU6371DLF2" Then
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Then
                                    
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
                
                 If rv0 = 1 Then
                  rv1 = 1  '----------- AU6371S3 dp not have CF slot
                 Else
                   rv1 = 4
                 End If
                 
               '  Call LabelMenu(1, rv1, rv0)
            
                 '     Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
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
                   rv2 = 4
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
                Call MsecDelay(AU6371EL_BootTime + 1.2)
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect Down Fail"
                    End
                 End If
                 
                  
                   ReaderExist = 0
                 
                ClosePipe
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  MS Card test
                '================================================
                  
                  If rv3 = 1 Then
                     rv4 = 1
                  Else
                    rv4 = 4
                  End If
                      
               
                   '  Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
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
                  Call MsecDelay(AU6371EL_BootTime + 1)
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect Down Fail"
                    End
                 End If
               
                ReaderExist = 0
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
Public Sub AU6430QLTestSub()
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
                
                 If Left(ChipName, 10) = "AU6430QLF2" Then
                  ChipName = "AU6371DLF20"
                 
                End If
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
                
                 If rv0 = 1 Then
                  rv1 = 1  '----------- AU6371S3 dp not have CF slot
                 Else
                   rv1 = 4
                 End If
                 
               '  Call LabelMenu(1, rv1, rv0)
            
                 '     Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
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
                   rv2 = 4
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
                  
                  If rv3 = 1 Then
                     rv4 = 1
                  Else
                    rv4 = 4
                  End If
                      
               
                   '  Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
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

Public Sub AU6430QLSortingSub()

Dim k As Long

'For k = 0 To 65535
'    Pattern_AU6377(k) = &H55
'Next k
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
                
                 If Left(ChipName, 10) = "AU6430QLS1" Then
                  ChipName = "AU6371DLF20"
                 
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
                 
     
            
                 rv0 = 1
                 rv1 = 1
                 rv2 = 1
                 rv3 = 1
      
                  ReaderExist = 0
                
           
                 rv4 = 1
                If rv4 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                End If
                
                  Call MsecDelay(AU6371EL_BootTime * 2 + 1.2)
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect Down Fail"
                    End
                 End If
                 If ChipName = "AU6371EL" Then
                   ReaderExist = 0
                 End If
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, "058f")
                 ClosePipe
                If rv5 = 1 Then
                OldLBa = LBA
                  For i = 1 To 800
                     OpenPipe
                        LBA = LBA + 1
                        rv5 = Write_Data_AU6377(LBA, 0, 65536)
                    ClosePipe
                    If rv5 <> 1 Then
                     '  Exit For
                    End If
                  Next i
                   
                    End If
                    Tester.Print i
                     Call MsecDelay(3)
                      ReaderExist = 0
                ClosePipe
                rv5 = CBWTest_New(0, 1, "058f")
                If i > 0 Then
                    If rv5 <> 0 Then
                      rv5 = TestUnitSpeed(0)
        
                        If rv5 = 0 Then
                           
                           rv5 = 2
                           UsbSpeedTestResult = 2
                          
                        End If
                    End If
                    
                 Else
                    If i <> 401 Then
                      rv5 = 2
                    End If
               End If
                 
                LBA = OldLBa
                
                
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
Public Function CBWTest_New_AU6430Sorting(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

   CBWDataTransferLength = 512
 
 

    If PreSlotStatus <> 1 Then
        CBWTest_New_AU6430Sorting = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New_AU6430Sorting = 0
    If LBA > 25 * 1024 Then
        LBA = 0
    End If
    '========================================
     TmpString = ""
    If ReaderExist = 0 Then
        Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName(Vid_PID)
        Loop While TmpString = "" And TimerCounter < 10
    End If
    '=======================================
    If ReaderExist = 0 And TmpString <> "" Then
      ReaderExist = 1
    End If
    '=======================================
    If ReaderExist = 0 And TmpString = "" Then
      CBWTest_New_AU6430Sorting = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_New_AU6430Sorting = 2   ' Write fail
      Exit Function
    End If
 
    '======================================
    
    
     ' for unitSpeed
    
     TmpInteger = TestUnitSpeed(Lun)
    
     If TmpInteger = 0 Then
        
        CBWTest_New_AU6430Sorting = 2   ' usb 2.0 high speed fail
        UsbSpeedTestResult = 2
        Exit Function
     End If
    
    
    
    TmpInteger = TestUnitReady(Lun)
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
           CBWTest_New_AU6430Sorting = 2  'Write fail
           Exit Function
        End If
        
    End If
    '======================================
  '  If ChipName = "AU6371" Or ChipName = "AU6371S3" Then
  '      TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
  '  End If
    
    TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
    
     TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
         CBWTest_New_AU6430Sorting = 2  'write fail
        '  Exit Function
     End If
    
      
   ' TmpInteger = Write_Data(LBA, Lun, CBWDataTransferLength)
     
   ' If TmpInteger = 0 Then
   '     CBWTest_New_AU6430Sorting = 2  'write fail
   '     Exit Function
   ' End If
    
    TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_AU6430Sorting = 3    'Read fail
        Exit Function
    End If
     
   ' For i = 0 To CBWDataTransferLength - 1
   
   '     If ReadData(i) <> Pattern(i) Then
   '       CBWTest_New_AU6430Sorting = 3    'Read fail
   '       Exit Function
   '     End If
    
   ' Next
    
     
    
    
    
    CBWTest_New_AU6430Sorting = 1
        
    
    End Function
    
 Public Function CBWTest_New_AU6256XLS17(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

   CBWDataTransferLength = 512
 
 

    If PreSlotStatus <> 1 Then
        CBWTest_New_AU6256XLS17 = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New_AU6256XLS17 = 0
    If LBA > 25 * 1024 Then
        LBA = 0
    End If
    '========================================
     TmpString = ""
    If ReaderExist = 0 Then
        Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName(Vid_PID)
        Loop While TmpString = "" And TimerCounter < 10
    End If
    '=======================================
    If ReaderExist = 0 And TmpString <> "" Then
      ReaderExist = 1
    End If
    '=======================================
    If ReaderExist = 0 And TmpString = "" Then
      CBWTest_New_AU6256XLS17 = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_New_AU6256XLS17 = 2   ' Write fail
      Exit Function
    End If
 
    '======================================
    
    
     ' for unitSpeed
    
     TmpInteger = TestUnitSpeed(Lun)
    
     If TmpInteger = 0 Then
        
        CBWTest_New_AU6256XLS17 = 2   ' usb 2.0 high speed fail
        UsbSpeedTestResult = 2
        Exit Function
     End If
    
    
    
  
    
     
    
    
    
    CBWTest_New_AU6256XLS17 = 1
        
    
    End Function
    
    
    
    Public Function CBWTest_New_AU6256X2Sorting(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

   CBWDataTransferLength = 512
 
 

    If PreSlotStatus <> 1 Then
        CBWTest_New_AU6256X2Sorting = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New_AU6256X2Sorting = 0
    If LBA > 25 * 1024 Then
        LBA = 0
    End If
    '========================================
     TmpString = ""
    If ReaderExist = 0 Then
        Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceNameMulti(Vid_PID)
        Loop While TmpString = "" And TimerCounter < 10
    End If
    '=======================================
    If ReaderExist = 0 And TmpString <> "" Then
      ReaderExist = 1
    End If
    '=======================================
    If ReaderExist = 0 And TmpString = "" Then
      CBWTest_New_AU6256X2Sorting = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_New_AU6256X2Sorting = 2   ' Write fail
      Exit Function
    End If
 
    '======================================
    
    
     ' for unitSpeed
    
     TmpInteger = TestUnitSpeed(Lun)
    
     If TmpInteger = 0 Then
        
        CBWTest_New_AU6256X2Sorting = 2   ' usb 2.0 high speed fail
        UsbSpeedTestResult = 2
        Exit Function
     End If
    
    
    
    TmpInteger = TestUnitReady(Lun)
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
           CBWTest_New_AU6256X2Sorting = 2  'Write fail
           Exit Function
        End If
        
    End If
    '======================================
  '  If ChipName = "AU6371" Or ChipName = "AU6371S3" Then
  '      TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
  '  End If
    
    TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
    
     TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
         CBWTest_New_AU6256X2Sorting = 2  'write fail
        '  Exit Function
     End If
    
      
   ' TmpInteger = Write_Data(LBA, Lun, CBWDataTransferLength)
     
   ' If TmpInteger = 0 Then
   '     CBWTest_New_AU6256X2Sorting = 2  'write fail
   '     Exit Function
   ' End If
    
    TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_AU6256X2Sorting = 3    'Read fail
        Exit Function
    End If
     
   ' For i = 0 To CBWDataTransferLength - 1
   
   '     If ReadData(i) <> Pattern(i) Then
   '       CBWTest_New_AU6256X2Sorting = 3    'Read fail
   '       Exit Function
   '     End If
    
   ' Next
    
     
    
    
    
    CBWTest_New_AU6256X2Sorting = 1
        
    
    End Function
    
 Public Function CBWTest_New_AU6256X2_1Sorting(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

   CBWDataTransferLength = 2048
 
 

    If PreSlotStatus <> 1 Then
        CBWTest_New_AU6256X2_1Sorting = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New_AU6256X2_1Sorting = 0
    If LBA > 25 * 1024 Then
        LBA = 0
    End If
    '========================================
     TmpString = ""
    If ReaderExist = 0 Then
        Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceNameMulti(Vid_PID)
        Loop While TmpString = "" And TimerCounter < 10
    End If
    '=======================================
    If ReaderExist = 0 And TmpString <> "" Then
      ReaderExist = 1
    End If
    '=======================================
    If ReaderExist = 0 And TmpString = "" Then
      CBWTest_New_AU6256X2_1Sorting = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_New_AU6256X2_1Sorting = 2   ' Write fail
      Exit Function
    End If
 
    '======================================
    
    
     ' for unitSpeed
    
     TmpInteger = TestUnitSpeed(Lun)
    
     If TmpInteger = 0 Then
        
        CBWTest_New_AU6256X2_1Sorting = 2   ' usb 2.0 high speed fail
        UsbSpeedTestResult = 2
        Exit Function
     End If
    
    
    
    TmpInteger = TestUnitReady(Lun)
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
           CBWTest_New_AU6256X2_1Sorting = 2  'Write fail
           Exit Function
        End If
        
    End If
    '======================================
  '  If ChipName = "AU6371" Or ChipName = "AU6371S3" Then
  '      TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
  '  End If
    
    TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
    
     TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
         CBWTest_New_AU6256X2_1Sorting = 2  'write fail
        '  Exit Function
     End If
    
      
     TmpInteger = Write_Data(LBA, Lun, CBWDataTransferLength)
     
     If TmpInteger = 0 Then
         CBWTest_New_AU6256X2_1Sorting = 2  'write fail
         Exit Function
     End If
    
    TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_AU6256X2_1Sorting = 3    'Read fail
        Exit Function
    End If
     
     For i = 0 To CBWDataTransferLength - 1
   
         If ReadData(i) <> Pattern(i) Then
           CBWTest_New_AU6256X2_1Sorting = 3    'Read fail
           Exit Function
         End If
    
     Next
    
     
    
    
    
    CBWTest_New_AU6256X2_1Sorting = 1
        
    
    End Function
    
Public Function Read_DataAU6430Sorting(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte
Dim CBW(0 To 30) As Byte
Dim NumberOfBytesWritten As Long
Dim CBWDataTransferLen(0 To 3) As Byte
  
Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim i As Long
Dim tmpV(0 To 2) As Long
Dim opcode As Byte

Dim CSW(0 To 12) As Byte

Dim NumberOfBytesRead As Long

For i = 0 To 30
   
        CBW(i) = 0
    
Next i

For i = 0 To CBWDataTransferLength
ReadData(i) = 0
Next

Const CBWSignature_0 = &H55
Const CBWSignature_1 = &H53
Const CBWSignature_2 = &H42
Const CBWSignature_3 = &H43


Const CBWTag_0 = &H1
Const CBWTag_1 = &H2
Const CBWTag_2 = &H3
Const CBWTag_3 = &H4


'/////////////////// CBW signature

CBW(0) = CBWSignature_0
CBW(1) = CBWSignature_1
CBW(2) = CBWSignature_2
CBW(3) = CBWSignature_3

'/////////////////  CBW Tag

CBW(4) = CBWTag_0
CBW(5) = CBWTag_1
CBW(6) = CBWTag_2
CBW(7) = CBWTag_3

CBWDataTransferLen(0) = (CBWDataTransferLength Mod 256)
tmpV(0) = Int(CBWDataTransferLength / 256)
CBWDataTransferLen(1) = (tmpV(0) Mod 256)
tmpV(1) = Int(tmpV(0) / 256)
CBWDataTransferLen(2) = (tmpV(1) Mod 256)
tmpV(2) = Int((tmpV(1) / 256))
CBWDataTransferLen(3) = (tmpV(2) Mod 256)

CBW(8) = CBWDataTransferLen(0)  '00
CBW(9) = CBWDataTransferLen(1)  '08
CBW(10) = CBWDataTransferLen(2) '00
CBW(11) = CBWDataTransferLen(3) '00

'///////////////  CBW Flag
CBW(12) = &H80                 '80

'////////////// LUN
CBW(13) = Lun                    '00

'///////////// CBD Len
CBW(14) = &HA                '0a

'////////////  UFI command

CBW(15) = &H28
CBW(16) = Lun * 32
LBAByte(0) = (LBA Mod 256)
tmpV(0) = Int(LBA / 256)
LBAByte(1) = (tmpV(0) Mod 256)
tmpV(1) = Int(tmpV(0) / 256)
LBAByte(2) = (tmpV(1) Mod 256)
tmpV(2) = Int((tmpV(1) / 256))
LBAByte(3) = (tmpV(2) Mod 256)

CBW(17) = LBAByte(3)         '00
CBW(18) = LBAByte(2)         '00
CBW(19) = LBAByte(1)         '00
CBW(20) = LBAByte(0)         '40

'/////////////  Reverve
CBW(21) = 0

'//////////// Transfer Len

TransferLen = Int(CBWDataTransferLength / 512)

TransferLenLSB = (TransferLen Mod 256)
tmpV(0) = Int(TransferLen / 256)
TransferLenMSB = (tmpV(0) / 256)

CBW(22) = TransferLenMSB      '00
CBW(23) = TransferLenLSB      '04

For i = 24 To 30
    CBW(i) = 0
Next

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
 
Dim result As Long

'1. CBW command

 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)    'out

If result = 0 Then
 Read_DataAU6430Sorting = 0
 Exit Function
End If

'2. Readdata stage
 
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
          
If result = 0 Or NumberOfBytesRead <> 512 Then
 Read_DataAU6430Sorting = 0
 Exit Function
End If
 
 

'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
  
 
 
If result = 0 Or NumberOfBytesRead <> 13 Then
 Read_DataAU6430Sorting = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
    Read_DataAU6430Sorting = 0
Else
     Read_DataAU6430Sorting = 1
   
End If

 
End Function
Public Sub AU6430QLS12SortingSub()

Dim k As Long
Dim OldTime
Dim CycleTime
Dim TestCycle As Integer
TestCycle = 100
Dim i As Integer
Dim Vol As Single
Tester.Cls
                
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
              
                 
       
                 PowerSet (1345)
                 
                
                  Call MsecDelay(AU6371EL_BootTime * 2 + 1.2)
             
              'Initial test
              
                 ClosePipe
                    rv0 = CBWTest_New_AU6430Sorting(0, 1, "058f")
                 ClosePipe
                 
                  Tester.Print rv0, " \\initial test"
                  Call LabelMenu(0, rv0, 1)
             ' Sorting Test
                 
                If rv0 = 1 Then
                    OldLBa = LBA
                    For Vol = 3.45 To 3.095 Step -0.01
                    
                    Call GPIBWrite("VSET1 " & CStr(Vol))
                    DoEvents
                    For i = 1 To TestCycle
                        OldTime = Timer
                        OpenPipe
                        LBA = LBA + 1
                            
                             rv1 = Read_DataAU6430Sorting(LBA, 0, 512)
                        ClosePipe
                        CycleTime = Timer - OldTimer
                         
                        If rv1 <> 1 And Abs(CycleTime) > 3 Then
                           Tester.Print "fail Clycle"; i
                        
                            Exit For
                        End If
                    Next i
                    Tester.Print "Vol="; Vol; " RV="; rv1
                    If rv1 <> 1 Then
                       rv1 = 2
                        Exit For
                    End If
                    Next Vol
                     
                    
                   
                End If
                
                
                 If rv1 <> 1 Then
                       rv1 = 2
                        
                 End If
                 Tester.Print rv1, " \\ 3.45V ~ 3.1 V cycle read  test"
                Call LabelMenu(1, rv1, rv0)
              ' test speed
                
              
               If rv1 = 1 Then
                    Tester.Cls
                    OldLBa = LBA
                    For Vol = 3.09 To 3.005 Step -0.01
                    
                    Call GPIBWrite("VSET1 " & CStr(Vol))
                    DoEvents
                    For i = 1 To TestCycle
                        OldTime = Timer
                        OpenPipe
                        LBA = LBA + 1
                            
                            rv2 = Read_DataAU6430Sorting(LBA, 0, 512)
                        ClosePipe
                        CycleTime = Timer - OldTimer
                        If rv2 <> 1 And Abs(CycleTime) > 3 Then
                           Tester.Print "fail Clycle"; i
                        
                            Exit For
                        End If
                    Next i
                    Tester.Print "Vol2="; Vol; " RV2="; rv2
                    If rv2 <> 1 Then
                       rv2 = 2
                        Exit For
                    End If
                    Next Vol
                     
                    
                   
                End If
                
                  If rv2 <> 1 Then
                       rv2 = 2
                        
                 End If
                 Tester.Print rv2, " \\3.09 V ~ 3.01 V cycle read  test"
                Call LabelMenu(2, rv2, rv1)
              
              
              
              
              
              If rv2 = 1 Then
                Call MsecDelay(3)
                      ReaderExist = 0
                ClosePipe
                  rv3 = CBWTest_New_AU6430Sorting(0, 1, "058f")
                If i > 0 Then
                    If rv3 <> 0 Then
                        rv3 = TestUnitSpeed(0)
        
                        If rv3 = 0 Then
                           
                           rv3 = 2
                           UsbSpeedTestResult = 2
                          
                        End If
                    End If
                    
                 Else
                    If i <> TestCycle + 1 Then
                      rv3 = 2
                    End If
               End If
               
               If rv3 <> 1 Then
                rv3 = 2
               End If
                 
               
                
                ClosePipe
                Call LabelMenu(3, rv3, rv2)
                Tester.Print rv3, " \\Final Speed check"
              End If
              
               LBA = OldLBa
                
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
                       
                            
                        ElseIf rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub

Public Sub AU6430QLS13SortingSub()

Dim k As Long
Dim OldTime
Dim CycleTime
Dim TestCycle As Integer
TestCycle = 100
Dim i As Integer
Dim Vol As Single
Tester.Cls
                
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
              
                 
       
                 PowerSet (1345)
                 
                
                  Call MsecDelay(AU6371EL_BootTime * 2 + 1.2)
             
              'Initial test
              
                 ClosePipe
                    rv0 = CBWTest_New_AU6430Sorting(0, 1, "058f")
                 ClosePipe
                 
                  Tester.Print rv0, " \\initial test"
                  Call LabelMenu(0, rv0, 1)
             ' Sorting Test
               
                If rv0 = 1 Then
                    OldLBa = LBA
                    For Vol = 3.45 To 3.135 Step -0.01
                    
                    Call GPIBWrite("VSET1 " & CStr(Vol))
                    DoEvents
                    For i = 1 To TestCycle
                        OldTime = Timer
                        OpenPipe
                        LBA = LBA + 1
                            
                             rv1 = Read_DataAU6430Sorting(LBA, 0, 512)
                        ClosePipe
                        CycleTime = Timer - OldTimer
                         
                        If rv1 <> 1 And Abs(CycleTime) > 3 Then
                           Tester.Print "fail Clycle"; i
                        
                            Exit For
                        End If
                    Next i
                    Tester.Print "Vol="; Vol; " RV="; rv1
                    If rv1 <> 1 Then
                       rv1 = 2
                        Exit For
                    End If
                    Next Vol
                     
                    
                   
                End If
                
                
                 If rv1 <> 1 Then
                       rv1 = 2
                        
                 End If
                 Tester.Print rv1, " \\ 3.45V ~ 3.14 V cycle read  test"
                Call LabelMenu(1, rv1, rv0)
              ' test speed
                
              
               If rv1 = 1 Then
                    Tester.Cls
                    OldLBa = LBA
                    For Vol = 3.13 To 3.095 Step -0.01
                    
                    Call GPIBWrite("VSET1 " & CStr(Vol))
                    DoEvents
                    For i = 1 To TestCycle
                        OldTime = Timer
                        OpenPipe
                        LBA = LBA + 1
                            
                            rv2 = Read_DataAU6430Sorting(LBA, 0, 512)
                        ClosePipe
                        CycleTime = Timer - OldTimer
                        If rv2 <> 1 And Abs(CycleTime) > 3 Then
                           Tester.Print "fail Clycle"; i
                        
                            Exit For
                        End If
                    Next i
                    Tester.Print "Vol2="; Vol; " RV2="; rv2
                    If rv2 <> 1 Then
                       rv2 = 2
                        Exit For
                    End If
                    Next Vol
                     
                    
                   
                End If
                
                  If rv2 <> 1 Then
                       rv2 = 2
                        
                 End If
                 Tester.Print rv2, " \\3.13 V ~ 3.10 V cycle read  test"
                Call LabelMenu(2, rv2, rv1)
              
                
              '============================================================
              
               If rv2 = 1 Then
                   ' Tester.Cls
                    OldLBa = LBA
                    For Vol = 3.09 To 3.005 Step -0.01
                    
                    Call GPIBWrite("VSET1 " & CStr(Vol))
                    DoEvents
                    For i = 1 To TestCycle
                        OldTime = Timer
                        OpenPipe
                        LBA = LBA + 1
                            
                            rv3 = Read_DataAU6430Sorting(LBA, 0, 512)
                        ClosePipe
                        CycleTime = Timer - OldTimer
                        If rv3 <> 1 And Abs(CycleTime) > 3 Then
                           Tester.Print "fail Clycle"; i
                        
                            Exit For
                        End If
                    Next i
                    Tester.Print "Vol3="; Vol; " RV3="; rv3
                    If rv3 <> 1 Then
                       rv3 = 2
                        Exit For
                    End If
                    Next Vol
                     
                    
                   
                End If
                
                  If rv3 <> 1 Then
                       rv3 = 2
                        
                 End If
                 Tester.Print rv3, " \\3.09 V ~ 3.01 V cycle read  test"
                Call LabelMenu(3, rv2, rv1)
              
           
              
              If rv3 = 1 Then
                Call MsecDelay(3)
                      ReaderExist = 0
                ClosePipe
                  rv4 = CBWTest_New_AU6430Sorting(0, 1, "058f")
                If i > 0 Then
                    If rv4 <> 0 Then
                        rv4 = TestUnitSpeed(0)
        
                        If rv4 = 0 Then
                           
                           rv4 = 2
                           UsbSpeedTestResult = 2
                          
                        End If
                    End If
                    
                 Else
                    If i <> TestCycle + 1 Then
                      rv4 = 2
                    End If
               End If
         
               
                
                ClosePipe
                Call LabelMenu(4, rv4, rv3)
                Tester.Print rv4, " \\Final Speed check"
              End If
               
              If rv3 = 1 Then
               If rv4 <> 1 Then
                rv0 = 0
               End If
              End If
               
               
               LBA = OldLBa
                
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
                     
                         ElseIf rv2 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv2 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                       
                        ElseIf rv3 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv3 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                        ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub
Public Sub AU698XHLS20SortingSub()

Dim k As Long
Dim OldTime
Dim CycleTime
Dim TestCycle As Integer
TestCycle = 100
Dim i As Integer
Dim Vol As Single
Tester.Cls
               Tester.Print "Add current Clamp at 0.15A"

              If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                 End If
                
                
                 '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                 CardResult = DO_WritePort(card, Channel_P1A, &H0)  'Power Enable
                 Call MsecDelay(1.8)    'power on time
                 

                
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
              
                 
       
                 PowerSet (1525)
                 
                
                  Call MsecDelay(AU6371EL_BootTime * 2 + 1.2)
             
              'Initial test
              
                 ClosePipe
                    rv0 = CBWTest_New_AU6430Sorting(0, 1, "058f")
                 ClosePipe
                 
                  Tester.Print rv0, " \\initial test"
                  Call LabelMenu(0, rv0, 1)
             ' Sorting Test
               
               Dim vs As Single
               
                If rv0 = 1 Then
                    OldLBa = LBA
                    For Vol = 525 To 474 Step -5
                    vs = Vol / 100
                    Call GPIBWrite("VSET1 " & vs)
                    DoEvents
                    For i = 1 To TestCycle
                        OldTime = Timer
                        OpenPipe
                        LBA = LBA + 1
                            
                             rv1 = Read_DataAU6430Sorting(LBA, 0, 512)
                        ClosePipe
                        CycleTime = Timer - OldTimer
                         
                        If rv1 <> 1 And Abs(CycleTime) > 3 Then
                           Tester.Print "fail Clycle"; i
                        
                            Exit For
                        End If
                    Next i
                    Tester.Print "Vol="; vs; " RV="; rv1
                    If rv1 <> 1 Then
                       rv1 = 2
                        Exit For
                    End If
                    Next Vol
                     
                    
                   
                End If
                
                
                 If rv1 <> 1 Then
                       rv1 = 2
                        
                 End If
                 Tester.Print rv1, " \\ 5.25V ~ 4.75 V cycle read  test"
                Call LabelMenu(1, rv1, rv0)
              ' test speed
                
              
                rv2 = 1
                rv3 = 1
              
                
              '============================================================
             
           
              
              If rv1 = 1 Then
                Call MsecDelay(3)
                      ReaderExist = 0
                ClosePipe
                  rv4 = CBWTest_New_AU6430Sorting(0, 1, "058f")
                If i > 0 Then
                    If rv4 <> 0 Then
                        rv4 = TestUnitSpeed(0)
        
                        If rv4 = 0 Then
                           
                           rv4 = 2
                           UsbSpeedTestResult = 2
                          
                        End If
                    End If
                    
                 Else
                    If i <> TestCycle + 1 Then
                      rv4 = 2
                    End If
               End If
         
               
                
                ClosePipe
                Call LabelMenu(4, rv4, rv3)
                Tester.Print rv4, " \\Final Speed check"
              End If
               
              If rv1 = 1 Then
               If rv4 <> 1 Then
                rv0 = 0
               End If
              End If
               
               
               LBA = OldLBa
                
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
                     
                         ElseIf rv2 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv2 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                       
                        ElseIf rv3 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv3 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                        ElseIf rv4 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub

Public Sub AU6337BLS10SortingSub()

Dim k As Long
Dim OldTime
Dim CycleTime
Dim TestCycle As Integer
TestCycle = 100
Dim i As Integer
Dim Vol As Single
Tester.Cls
             '  Tester.Print "Add current Clamp at 0.15A"

              If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                 End If
                
                
                 '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                 CardResult = DO_WritePort(card, Channel_P1A, &H0)  'Power Enable
                 Call MsecDelay(1.8)    'power on time
                 

                
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
              
                 
       
                 PowerSet (1530)
                 
                
                  Call MsecDelay(AU6371EL_BootTime * 2 + 1.2)
             
              'Initial test
              
                 ClosePipe
                    rv0 = CBWTest_New_AU6430Sorting(0, 1, "058f")
                 ClosePipe
                 
                  Tester.Print rv0, " \\initial test"
                  Call LabelMenu(0, rv0, 1)
             ' Sorting Test
               
               Dim vs As Single
               
                If rv0 = 1 Then
                    OldLBa = LBA
                    For Vol = 530 To 470 Step -1
                    
                    If Vol = 500 Then
                      Tester.Cls
                    End If
                    
                    vs = Vol / 100
                    Call GPIBWrite("VSET1 " & vs)
                    DoEvents
                    For i = 1 To TestCycle
                        OldTime = Timer
                        OpenPipe
                        LBA = LBA + 1
                            
                             rv1 = Read_DataAU6430Sorting(LBA, 0, 512)
                        ClosePipe
                        CycleTime = Timer - OldTimer
                         
                         
                        If rv1 <> 1 And Abs(CycleTime) > 3 Then
                           Tester.Print "fail Clycle"; i
                        
                            Exit For
                        End If
                    Next i
                    
                    
                    Tester.Print "Vol="; vs; " RV="; rv1
                    If rv1 <> 1 Then
                       rv1 = 2
                        Exit For
                    End If
                    Next Vol
                     
                    
                   
                End If
                
                
                 If rv1 <> 1 Then
                       rv1 = 2
                        
                 End If
                 Tester.Print rv1, " \\ 5.25V ~ 4.75 V cycle read  test"
                Call LabelMenu(1, rv1, rv0)
              ' test speed
                
              
                rv2 = 1
                rv3 = 1
              
                
              '============================================================
             
           
              
              If rv1 = 1 Then
                Call MsecDelay(3)
                      ReaderExist = 0
                ClosePipe
                  rv4 = CBWTest_New_AU6430Sorting(0, 1, "058f")
                If i > 0 Then
                    If rv4 <> 0 Then
                        rv4 = TestUnitSpeed(0)
        
                        If rv4 = 0 Then
                           
                           rv4 = 2
                           UsbSpeedTestResult = 2
                          
                        End If
                    End If
                    
                 Else
                    If i <> TestCycle + 1 Then
                      rv4 = 2
                    End If
               End If
         
               
                
                ClosePipe
                Call LabelMenu(4, rv4, rv3)
                Tester.Print rv4, " \\Final Speed check"
              End If
               
              If rv1 = 1 Then
               If rv4 <> 1 Then
                rv0 = 0
               End If
              End If
               
               
               LBA = OldLBa
                
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
                     
                         ElseIf rv2 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv2 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                       
                        ElseIf rv3 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv3 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                        ElseIf rv4 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub
Public Sub AU6430QLS10SortingSub()

Dim k As Long
Dim OldTime
Dim CycleTime
Dim TestCycle As Integer
TestCycle = 100
Dim i As Integer
Dim Vol As Single
Tester.Cls
                
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
              
                 
       
                 PowerSet (34)
                 
                
                  Call MsecDelay(AU6371EL_BootTime * 2 + 1.2)
             
              'Initial test
              
                 ClosePipe
                    rv0 = CBWTest_New_AU6430Sorting(0, 1, "058f")
                 ClosePipe
                 
                  Tester.Print rv0, " \\initial test"
                  Call LabelMenu(0, rv0, 1)
             ' Sorting Test
                 
                If rv0 = 1 Then
                    OldLBa = LBA
                    For Vol = 3.4 To 3.095 Step -0.01
                    
                    Call GPIBWrite("VSET1 " & CStr(Vol))
                    DoEvents
                    For i = 1 To TestCycle
                        OldTime = Timer
                        OpenPipe
                        LBA = LBA + 1
                            
                            rv1 = Read_DataAU6430Sorting(LBA, 0, 512)
                        ClosePipe
                        CycleTime = Timer - OldTimer
                        If rv1 <> 1 And Abs(CycleTime) > 3 Then
                           Tester.Print "fail Clycle"; i
                        
                            Exit For
                        End If
                    Next i
                    Tester.Print "Vol="; Vol; " RV="; rv1
                    If rv1 <> 1 Then
                       rv1 = 2
                        Exit For
                    End If
                    Next Vol
                     
                    
                   
                End If
                
                
                 If rv1 <> 1 Then
                       rv1 = 2
                        
                 End If
                 Tester.Print rv1, " \\ 3.4V ~ 3.1 V cycle read  test"
                Call LabelMenu(1, rv1, rv0)
              ' test speed
                
              
               If rv1 = 1 Then
                    OldLBa = LBA
                    For Vol = 3.09 To 3.005 Step -0.01
                    
                    Call GPIBWrite("VSET1 " & CStr(Vol))
                    DoEvents
                    For i = 1 To TestCycle
                        OldTime = Timer
                        OpenPipe
                        LBA = LBA + 1
                            
                            rv2 = Read_DataAU6430Sorting(LBA, 0, 512)
                        ClosePipe
                        CycleTime = Timer - OldTimer
                        If rv2 <> 1 And Abs(CycleTime) > 3 Then
                           Tester.Print "fail Clycle"; i
                        
                            Exit For
                        End If
                    Next i
                    Tester.Print "Vol2="; Vol; " RV2="; rv1
                    If rv2 <> 1 Then
                       rv2 = 2
                        Exit For
                    End If
                    Next Vol
                     
                    
                   
                End If
                
                  If rv2 <> 1 Then
                       rv2 = 2
                        
                 End If
                 Tester.Print rv2, " \\3.09 V ~ 3.01 V cycle read  test"
                Call LabelMenu(2, rv2, rv1)
              
              
              
              
              
              If rv2 = 1 Then
                Call MsecDelay(3)
                      ReaderExist = 0
                ClosePipe
                  rv3 = CBWTest_New_AU6430Sorting(0, 1, "058f")
                If i > 0 Then
                    If rv3 <> 0 Then
                        rv3 = TestUnitSpeed(0)
        
                        If rv3 = 0 Then
                           
                           rv3 = 2
                           UsbSpeedTestResult = 2
                          
                        End If
                    End If
                    
                 Else
                    If i <> TestCycle + 1 Then
                      rv3 = 2
                    End If
               End If
               
               If rv3 <> 1 Then
                rv3 = 2
               End If
                 
               
                
                ClosePipe
                Call LabelMenu(3, rv3, rv2)
                Tester.Print rv3, " \\Final Speed check"
              End If
              
               LBA = OldLBa
                
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
                       
                            
                        ElseIf rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub
Public Sub AU6433DFS10SortingSub()

Dim k As Long
Dim OldTime
Dim CycleTime
Dim TestCycle As Integer
TestCycle = 100
Dim i As Integer
Dim Vol As Single
Tester.Cls
                    If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
              
   
                
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
                 PowerSet (1)
                 
                   CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                  Call MsecDelay(AU6371EL_BootTime * 2 + 1.2)
             
              'Initial test
              
                 ClosePipe
                    rv0 = CBWTest_New_AU6430Sorting(0, 1, "058f")
                 ClosePipe
                 
                  Tester.Print rv0, " \\initial test"
                  Call LabelMenu(0, rv0, 1)
             ' Sorting Test
                 
                If rv0 = 1 Then
                    OldLBa = LBA
                    For Vol = 3.6 To 2.995 Step -0.01
                    
                    Call GPIBWrite("VSET1 " & CStr(Vol))
                    DoEvents
                    For i = 1 To TestCycle
                        OldTime = Timer
                        OpenPipe
                        LBA = LBA + 1
                            
                            rv1 = Read_DataAU6430Sorting(LBA, 0, 512)
                        ClosePipe
                        CycleTime = Timer - OldTimer
                        If rv1 <> 1 And Abs(CycleTime) > 3 Then
                           Tester.Print "fail Clycle"; i
                        
                            Exit For
                        End If
                    Next i
                    Tester.Print "Vol="; Vol; " RV="; rv1
                    If rv1 <> 1 Then
                       rv1 = 2
                        Exit For
                    End If
                    Next Vol
                     
                    
                   
                End If
                
                
                 If rv1 <> 1 Then
                       rv1 = 2
                        
                 End If
                 Tester.Print rv1, " \\ 3.6V ~ 3.0 V cycle read  test"
                Call LabelMenu(1, rv1, rv0)
              ' test speed
                
              
              
              
              
              If rv2 = 1 Then
                Call MsecDelay(3)
                      ReaderExist = 0
                ClosePipe
                  rv3 = CBWTest_New_AU6430Sorting(0, 1, "058f")
                If i > 0 Then
                    If rv3 <> 0 Then
                        rv3 = TestUnitSpeed(0)
        
                        If rv3 = 0 Then
                           
                           rv3 = 2
                           UsbSpeedTestResult = 2
                          
                        End If
                    End If
                    
                 Else
                    If i <> TestCycle + 1 Then
                      rv3 = 2
                    End If
               End If
               
               If rv3 <> 1 Then
                rv3 = 2
               End If
                 
               
                
                ClosePipe
                Call LabelMenu(3, rv3, rv2)
                Tester.Print rv3, " \\Final Speed check"
              End If
              
               LBA = OldLBa
                
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
                       
                            
                        ElseIf rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub

Public Sub AU6471FLS10SortingSub()

Dim k As Long
Dim OldTime
Dim CycleTime
Dim TestCycle As Integer
TestCycle = 100
Dim i As Integer
Dim Vol As Single
Tester.Cls
              If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
              
   
                
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
                 PowerSet (1)
                 
                   CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                  Call MsecDelay(AU6371EL_BootTime * 2 + 1.2)
             
              'Initial test
              
                 ClosePipe
                    rv0 = CBWTest_New_AU6430Sorting(0, 1, "1984")
                 ClosePipe
                 
                  Tester.Print rv0, " \\initial test"
                  Call LabelMenu(0, rv0, 1)
             ' Sorting Test
                 
                If rv0 = 1 Then
                    OldLBa = LBA
                    For Vol = 3.6 To 3.305 Step -0.01
                    
                    Call GPIBWrite("VSET1 " & CStr(Vol))
                    DoEvents
                    For i = 1 To TestCycle
                        OldTime = Timer
                        OpenPipe
                        LBA = LBA + 1
                            
                            rv1 = Read_DataAU6430Sorting(LBA, 0, 512)
                        ClosePipe
                        CycleTime = Timer - OldTimer
                        If rv1 <> 1 And Abs(CycleTime) > 3 Then
                           Tester.Print "fail Clycle"; i
                        
                            Exit For
                        End If
                    Next i
                    Tester.Print "Vol="; Vol; " RV="; rv1
                    If rv1 <> 1 Then
                       rv1 = 2
                        Exit For
                    End If
                    Next Vol
                     
                    
                   
                End If
                
                
                 If rv1 <> 1 Then
                       rv1 = 2
                        
                 End If
                 Tester.Print rv1, " \\ 3.6V ~ 3.31 V cycle read  test"
                Call LabelMenu(1, rv1, rv0)
                  
                
                  If rv1 = 1 Then
                  Tester.Cls
                    OldLBa = LBA
                    For Vol = 3.3 To 2.995 Step -0.01
                    
                    Call GPIBWrite("VSET1 " & CStr(Vol))
                    DoEvents
                    For i = 1 To TestCycle
                        OldTime = Timer
                        OpenPipe
                        LBA = LBA + 1
                            
                            rv2 = Read_DataAU6430Sorting(LBA, 0, 512)
                        ClosePipe
                        CycleTime = Timer - OldTimer
                        If rv2 <> 1 And Abs(CycleTime) > 3 Then
                           Tester.Print "fail Clycle"; i
                        
                            Exit For
                        End If
                    Next i
                    Tester.Print "Vol2="; Vol; " RV2="; rv2
                    If rv2 <> 1 Then
                       rv2 = 2
                        Exit For
                    End If
                    Next Vol
                     
                    
                   
                End If
                
                  If rv2 <> 1 Then
                       rv2 = 2
                        
                 End If
                 Tester.Print rv2, " \\3.3 V ~ 3.0 V cycle read  test"
                Call LabelMenu(2, rv2, rv1)
              ' test speed
                
              
              
              
              
              If rv2 = 1 Then
                Call MsecDelay(3)
                      ReaderExist = 0
                ClosePipe
                  rv3 = CBWTest_New_AU6430Sorting(0, 1, "1984")
                If i > 0 Then
                    If rv3 <> 0 Then
                        rv3 = TestUnitSpeed(0)
        
                        If rv3 = 0 Then
                           
                           rv3 = 2
                           UsbSpeedTestResult = 2
                          
                        End If
                    End If
                    
                 Else
                    If i <> TestCycle + 1 Then
                      rv3 = 2
                    End If
               End If
               
               If rv3 <> 1 Then
                rv3 = 2
               End If
                 
               
                
                ClosePipe
                Call LabelMenu(3, rv3, rv2)
                Tester.Print rv3, " \\Final Speed check"
              End If
              
               LBA = OldLBa
                  CardResult = DO_WritePort(card, Channel_P1A, &H80)
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
                       
                            
                        ElseIf rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub
Public Sub AU6471FLS11SortingSub()

Dim k As Long
Dim OldTime
Dim CycleTime
Dim TestCycle As Integer
TestCycle = 100
Dim i As Integer
Dim Vol As Single
Tester.Cls
              If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
              
   
                
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
                 PowerSet (1)
                 
                   CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                  Call MsecDelay(AU6371EL_BootTime * 2 + 1.2)
             
              'Initial test
              
                 ClosePipe
                    rv0 = CBWTest_New_AU6430Sorting(0, 1, "058f")
                 ClosePipe
                 
                  Tester.Print rv0, " \\initial test"
                  Call LabelMenu(0, rv0, 1)
             ' Sorting Test
                 
                If rv0 = 1 Then
                    OldLBa = LBA
                    For Vol = 3.6 To 3.305 Step -0.01
                    
                    Call GPIBWrite("VSET1 " & CStr(Vol))
                    DoEvents
                    For i = 1 To TestCycle
                        OldTime = Timer
                        OpenPipe
                        LBA = LBA + 1
                            
                            rv1 = Read_DataAU6430Sorting(LBA, 0, 512)
                        ClosePipe
                        CycleTime = Timer - OldTimer
                        If rv1 <> 1 And Abs(CycleTime) > 3 Then
                           Tester.Print "fail Clycle"; i
                        
                            Exit For
                        End If
                    Next i
                    Tester.Print "Vol="; Vol; " RV="; rv1
                    If rv1 <> 1 Then
                       rv1 = 2
                        Exit For
                    End If
                    Next Vol
                     
                    
                   
                End If
                
                
                 If rv1 <> 1 Then
                       rv1 = 2
                        
                 End If
                 Tester.Print rv1, " \\ 3.6V ~ 3.31 V cycle read  test"
                Call LabelMenu(1, rv1, rv0)
                  
                
                  If rv1 = 1 Then
                  Tester.Cls
                    OldLBa = LBA
                    For Vol = 3.3 To 2.995 Step -0.01
                    
                    Call GPIBWrite("VSET1 " & CStr(Vol))
                    DoEvents
                    For i = 1 To TestCycle
                        OldTime = Timer
                        OpenPipe
                        LBA = LBA + 1
                            
                            rv2 = Read_DataAU6430Sorting(LBA, 0, 512)
                        ClosePipe
                        CycleTime = Timer - OldTimer
                        If rv2 <> 1 And Abs(CycleTime) > 3 Then
                           Tester.Print "fail Clycle"; i
                        
                            Exit For
                        End If
                    Next i
                    Tester.Print "Vol2="; Vol; " RV2="; rv2
                    If rv2 <> 1 Then
                       rv2 = 2
                        Exit For
                    End If
                    Next Vol
                     
                    
                   
                End If
                
                  If rv2 <> 1 Then
                       rv2 = 2
                        
                 End If
                 Tester.Print rv2, " \\3.3 V ~ 3.0 V cycle read  test"
                Call LabelMenu(2, rv2, rv1)
              ' test speed
                
              
              
              
              
              If rv2 = 1 Then
                Call MsecDelay(3)
                      ReaderExist = 0
                ClosePipe
                  rv3 = CBWTest_New_AU6430Sorting(0, 1, "058f")
                If i > 0 Then
                    If rv3 <> 0 Then
                        rv3 = TestUnitSpeed(0)
        
                        If rv3 = 0 Then
                           
                           rv3 = 2
                           UsbSpeedTestResult = 2
                          
                        End If
                    End If
                    
                 Else
                    If i <> TestCycle + 1 Then
                      rv3 = 2
                    End If
               End If
               
               If rv3 <> 1 Then
                rv3 = 2
               End If
                 
               
                
                ClosePipe
                Call LabelMenu(3, rv3, rv2)
                Tester.Print rv3, " \\Final Speed check"
              End If
              
               LBA = OldLBa
                  CardResult = DO_WritePort(card, Channel_P1A, &H80)
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
                       
                            
                        ElseIf rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub


Public Sub AU6256XLS15SortingSub()

Dim k As Long
Dim OldTime
Dim CycleTime
Dim TestCycle As Integer
TestCycle = 2500
Dim i As Integer
Dim Vol As Single
Tester.Cls
                 Call PowerSet2(1, "5.0", "0.5", 1, "5.0", "0.5", 1)
                 Call MsecDelay(0.8)
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
              
             
                
         '       Call MsecDelay(1.5)
                
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
                 Call MsecDelay(2#)
                
                 ClosePipe
                    rv0 = CBWTest_New_AU6430Sorting(0, 1, "1307")
                 ClosePipe
                 
                  Tester.Print rv0, " \\initial test"
                  Call LabelMenu(0, rv0, 1)
             ' Sorting Test
                 
                If rv0 = 1 Then
                    OldLBa = LBA
                    
                 
                    For i = 1 To TestCycle
                        OldTime = Timer
                        OpenPipe
                        LBA = LBA + 1
                            
                            rv1 = Read_DataAU6430Sorting(LBA, 0, 512)
                        ClosePipe
                        CycleTime = Timer - OldTimer
                        If rv1 <> 1 And Abs(CycleTime) > 3 Then
                           Tester.Print "fail Clycle"; i
                        
                            Exit For
                        End If
                    Next i
                    If rv1 <> 1 Then
                       rv1 = 2
                         
                    End If
                    
                     
                    
                   
                End If
                
                 Call PowerSet2(2, "5.0", "0.5", 1, "5.0", "0.5", 1)
                 Call MsecDelay(2.8)
                 If rv1 <> 1 Then
                       rv1 = 2
                        
                 End If
                 Tester.Print rv1, " \\ 3.6V ~ 3.31 V cycle read  test"
                Call LabelMenu(1, rv1, rv0)
                  
      
            
              
               LBA = OldLBa
                 
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
                       
                            
                        ElseIf rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub

Public Sub AU6256XLS14SortingSub()

'========================================================================
    '2010/5/6
    'Skip GPIB Power-control
'========================================================================
Dim k As Long
Dim OldTime
Dim CycleTime
Dim TestCycle As Integer
TestCycle = 2500
Dim i As Integer
Dim Vol As Single
Dim UsbSpeedTestResult As Integer
Dim XLS14_Flag As Integer
Tester.Cls

UsbSpeedTestResult = 0
XLS14_Flag = 0




                'Call PowerSet2(2, "5.0", "0.5", 1, "5.0", "0.5", 1)
                'Call MsecDelay(0.3)
                'Call PowerSet2(1, "5.0", "0.5", 1, "5.0", "0.5", 1)
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
SPEED_RT:
                

              
                CardResult = DO_WritePort(card, Channel_P1A, &HFE)
                fnScsi2usb2K_KillEXE
                MsecDelay (0.3)
                
                CardResult = DO_WritePort(card, Channel_P1A, &HFC)
                MsecDelay (0.5)
                
                 
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
                 Call MsecDelay(2#)
                
                 ClosePipe

                    rv0 = CBWTest_New_AU6430Sorting(0, 1, "1307")
                 ClosePipe
                 
                    If UsbSpeedTestResult <> 0 And XLS14_Flag = 0 Then      'RT speend error when first time fail
                       GoTo SPEED_RT
                       XLS21_Flag = 1
                    End If
                 
                  Tester.Print rv0, " \\initial test"
                  Call LabelMenu(0, rv0, 1)
             ' Sorting Test
                 
                If rv0 = 1 Then
                    OldLBa = LBA
                    
                    
                    For i = 1 To TestCycle
                        OldTime = Timer
                        OpenPipe
                        LBA = LBA + 1
                            
                            rv1 = Read_DataAU6430Sorting(LBA, 0, 512)
                        ClosePipe
                        CycleTime = Timer - OldTimer
                        If rv1 <> 1 And Abs(CycleTime) > 3 Then
                           Tester.Print "2nd fail Clycle"; i
                            Exit For
                        End If
                    Next i
                    
                    
                    If rv1 <> 1 Then
                       rv1 = 2
                         
                    End If
                    
                End If

                 If rv1 <> 1 Then
                       rv1 = 2
                        
                 End If
                 Tester.Print rv1, " \\ 3.6V ~ 3.31 V cycle read  test"
                Call LabelMenu(1, rv1, rv0)
                
                
           '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                MsecDelay (0.3)
      
            
              
               LBA = OldLBa
                 
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
                       
                            
                        ElseIf rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub
 
Public Sub AU6256XLS14SortingSub_Old()
On Error Resume Next
Dim k As Long
Dim OldTime
Dim CycleTime
Dim TestCycle As Integer
TestCycle = 2500
Dim i As Integer
Dim Vol As Single
Tester.Cls
                 Call PowerSet2(1, "5.0", "0.5", 1, "5.0", "0.5", 1)
                 Call MsecDelay(0.8)
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
              
             
                
         '       Call MsecDelay(1.5)
                
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
                 Call MsecDelay(2#)
                
                 ClosePipe
                    rv0 = CBWTest_New_AU6430Sorting(0, 1, "1307")
                 ClosePipe
                 
                  Tester.Print rv0, " \\initial test"
                  Call LabelMenu(0, rv0, 1)
             ' Sorting Test
                 
                If rv0 = 1 Then
                    OldLBa = LBA
                    
                 
                    For i = 1 To TestCycle
                        OldTime = Timer
                        OpenPipe
                        LBA = LBA + 1
                            
                            rv1 = Read_DataAU6430Sorting(LBA, 0, 512)
                        ClosePipe
                        CycleTime = Timer - OldTimer
                        If rv1 <> 1 And Abs(CycleTime) > 3 Then
                           Tester.Print "fail Clycle"; i
                        
                            Exit For
                        End If
                    Next i
                    If rv1 <> 1 Then
                       rv1 = 2
                         
                    End If
                    
                     
                    
                   
                End If
                
                 Call PowerSet2(2, "5.0", "0.5", 1, "5.0", "0.5", 1)
                 If rv1 <> 1 Then
                       rv1 = 2
                        
                 End If
                 Tester.Print rv1, " \\ 3.6V ~ 3.31 V cycle read  test"
                Call LabelMenu(1, rv1, rv0)
                  
      
            
              
               LBA = OldLBa
                 
AU6371DLResult:
                      If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                             AU6256Unknow = AU6256Unknow + 1
                           If AU6256Unknow > 2 Then
                             Shell "cmd /c shutdown -r  -t 0", vbHide
                           End If
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
                        ElseIf rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                             AU6256Unknow = 0
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub

Public Sub AU6256XLS19SortingSub()

Dim k As Long
Dim OldTime
Dim CycleTime
Dim TestCycle As Integer
TestCycle = 2500
Dim i As Integer
Dim Vol As Single
Tester.Cls
                 Call PowerSet2(1, "5.0", "0.5", 1, "5.0", "0.5", 1)
                 Call MsecDelay(0.8)
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
              
             
                
         '       Call MsecDelay(1.5)
                
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
                 Call MsecDelay(2#)
                 
                  HubPort = 1
                 ClosePipe
                    rv0 = CBWTest_New_AU6256X2_1Sorting(0, 1, "1307")
                 ClosePipe
                
                 
                 
                  Tester.Print rv0, " \\initial test"
                  Call LabelMenu(0, rv0, 1)
             ' Sorting Test
                 
                If rv0 = 1 Then
                    OldLBa = LBA
                    
                 
                    For i = 1 To TestCycle
                        OldTime = Timer
                        OpenPipe
                        LBA = LBA + 1
                            
                            rv1 = Read_DataAU6430Sorting(LBA, 0, 512)
                        ClosePipe
                        CycleTime = Timer - OldTimer
                        If rv1 <> 1 And Abs(CycleTime) > 3 Then
                           Tester.Print "fail Clycle"; i
                        
                            Exit For
                        End If
                    Next i
                    If rv1 <> 1 Then
                       rv1 = 2
                         
                    End If
                    
                     
                    
                   
                End If
                
                
                 
                
              
                 If rv1 <> 1 Then
                       rv1 = 2
                        
                 End If
                 Tester.Print rv1, " \\  cycle read  test"
                Call LabelMenu(1, rv1, rv0)
                    LBA = OldLBa
                  If rv1 = 1 Then
                           rv2 = 0
                            ReaderExist = 0
                            ClosePipe
                             rv2 = CBWTest_New(0, rv1, "9360")
                          ClosePipe
                          
                            Tester.Print rv2, " \\   AU9360 test"
                   End If
                     Call PowerSet2(2, "5.0", "0.5", 1, "5.0", "0.5", 1)
                Call LabelMenu(2, rv2, rv1)
                
                
                 
                 
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
                       
                            
                        ElseIf rv1 * rv0 * rv2 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub

Public Sub AU6256XLS1DSortingSub()
On Error Resume Next
Dim k As Long
Dim OldTime
Dim CycleTime
Dim TestCycle As Integer
TestCycle = 2500
Dim i As Integer
Dim Vol As Single
Dim TestAgainFlag As Byte
Dim TimeInterval
TestAgainFlag = 0
Tester.Cls
                 
                   If PCI7248InitFinish = 0 Then
                     PCI7248ExistAU6254
                   End If
                 
                     Call MsecDelay(1#)
                   CardResult = DO_WritePort(card, Channel_P1CH, &H0)
                   CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' open hub power
                 
                
              
SWITCH_TEST:
                 rv0 = 0
                 rv1 = 0
                 ReaderExist = 0
                 If TestAgainFlag = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1CH, &HF)
                     Call MsecDelay(0.2)
                      CardResult = DO_WritePort(card, Channel_P1A, &H0)
                     Call MsecDelay(1.2)
                     
                 Else
                   Call MsecDelay(0.8)
                 End If
                    
                   Call MsecDelay(1.5)
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
                 Call MsecDelay(2#)
                 
                  HubPort = 1
                
                 ClosePipe
                     rv0 = CBWTest_New_AU6256X2_1Sorting(0, 1, "1307")
                 ClosePipe
                
                 
                 
                  Tester.Print rv0, " \\initial test"
                  Call LabelMenu(0, rv0, 1)
             ' Sorting Test
             
                 If rv0 = 1 And TestAgainFlag = 1 Then
                 rv1 = 1
                    Else
                
                 rv1 = 4
               End If
                 
                If rv0 = 1 And TestAgainFlag = 0 Then
                    OldLBa = LBA
                    
                   Tester.Print " cycle test begin"
                    For i = 1 To TestCycle
                        OldTime = Timer
                        OpenPipe
                        LBA = LBA + 1
                            
                            rv1 = Read_DataAU6430Sorting(LBA, 0, 512)
                        ClosePipe
                        CycleTime = Timer - OldTimer
                        If rv1 <> 1 And Abs(CycleTime) > 3 Then
                           Tester.Print "fail Clycle"; i
                        
                            Exit For
                        End If
                    Next i
                    If rv1 <> 1 Then
                       rv1 = 2
                         
                    End If
                    
             
                End If
                
                
               
              
                 If rv1 <> 1 Then
                       rv1 = 2
                        
                 End If
                 Tester.Print rv1, " \\  cycle read  test"
                Call LabelMenu(1, rv1, rv0)
                
                '   CardResult = DO_WritePort(card, Channel_P1A, &HFA)
                 '  Call MsecDelay(2#)
                
                    LBA = OldLBa
                  If rv1 = 1 Then
                           rv2 = 0
                            ReaderExist = 0
                            ClosePipe
                             rv2 = CBWTest_New(0, rv1, "9360")
                          ClosePipe
                          
                            Tester.Print rv2, " \\   AU9360 test"
                   End If
                    
                Call LabelMenu(2, rv2, rv1)
                
                 If rv0 * rv1 * rv2 = PASS And TestAgainFlag = 0 Then
                   CardResult = DO_WritePort(card, Channel_P1A, &HFE)
                   TestAgainFlag = 1
                   
               
                  
                 
                   GoTo SWITCH_TEST
                   
                 End If
                
               
                    CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' open hub power
                   
               '  Call PowerSet2(2, "5.0", "0.5", 1, "5.0", "0.5", 1)
                
                 
                 
AU6256SortigResult:

                      CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' open hub power
                      Call MsecDelay(0.1)
                      If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                              AU6256Unknow = AU6256Unknow + 1
                           If AU6256Unknow > 2 Then
                             Shell "cmd /c shutdown -r  -t 0", vbHide
                           End If
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
                       
                            
                        ElseIf rv1 * rv0 * rv2 = PASS Then
                             TestResult = "PASS"
                             AU6256Unknow = 0
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub
Public Sub AU6256XLS1BSortingSub()
On Error Resume Next
Dim k As Long
Dim OldTime
Dim CycleTime
Dim TestCycle As Integer
TestCycle = 2500
Dim i As Integer
Dim Vol As Single
Dim TestAgainFlag As Byte
Dim TimeInterval
TestAgainFlag = 0
Tester.Cls
                   Call PowerSet2(1, "5.0", "0.5", 1, "5.0", "0.5", 1)
                   If PCI7248InitFinish = 0 Then
                     PCI7248ExistAU6254
                   End If
                 '  CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' open hub power
                     Call MsecDelay(1#)
                   CardResult = DO_WritePort(card, Channel_P1A, &HFE)  ' open hub power
                 
                 ' get Hub vid pid
                 
                  ReaderExist = 0
                  rv0 = 0
                  OldTimer = Timer
                  Tester.Cls
                  'Print "3"
                  Do
                      Call MsecDelay(0.2)
                      DoEvents
                       rv0 = AU6254_GetDevice(0, 1, "6254")
                       
                      TimeInterval = Timer - OldTimer
                  Loop While rv0 = 0 And TimeInterval < 15
                 
                  Tester.Print "rv0 ="; rv0; "TimeInterval="; TimeInterval
                  
                  If rv0 = 0 Then
                  GoTo AU6256SortigResult
                  
                  End If
              
SWITCH_TEST:
                 rv0 = 0
                 rv1 = 0
                 ReaderExist = 0
                 If TestAgainFlag = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1CH, &HF)
                     CardResult = DO_WritePort(card, Channel_P1A, &HFC)
                     Call MsecDelay(0.2)
                     CardResult = DO_WritePort(card, Channel_P1A, &HF8)
                    Call MsecDelay(1.2)
                 Else
                    CardResult = DO_WritePort(card, Channel_P1CH, &H0)
                     CardResult = DO_WritePort(card, Channel_P1A, &HFC)
                     Call MsecDelay(1#)
                     CardResult = DO_WritePort(card, Channel_P1A, &HF8)
                    Call MsecDelay(0.4)
                 End If
                 
                
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
              '   Call MsecDelay(2#)
                 
                  HubPort = 1
                '   HubPort = 0
                 ClosePipe
                     rv0 = CBWTest_New_AU6256X2_1Sorting(0, 1, "1307")
                    'rv0 = CBWTest_New(0, 1, "1307")
                 ClosePipe
                
                 
                 
                  Tester.Print rv0, " \\initial test"
                  Call LabelMenu(0, rv0, 1)
             ' Sorting Test
             
                 If rv0 = 1 And TestAgainFlag = 1 Then
                 rv1 = 1
                    Else
                
                 rv1 = 4
               End If
                 
                If rv0 = 1 And TestAgainFlag = 0 Then
                    OldLBa = LBA
                    
                   Tester.Print " cycle test begin"
                    For i = 1 To TestCycle
                        OldTime = Timer
                        OpenPipe
                        LBA = LBA + 1
                            
                            rv1 = Read_DataAU6430Sorting(LBA, 0, 512)
                        ClosePipe
                        CycleTime = Timer - OldTimer
                        If rv1 <> 1 And Abs(CycleTime) > 3 Then
                           Tester.Print "fail Clycle"; i
                        
                            Exit For
                        End If
                    Next i
                    If rv1 <> 1 Then
                       rv1 = 2
                         
                    End If
                    
                     
                    
                   
                End If
                
                
               
              
                 If rv1 <> 1 Then
                       rv1 = 2
                        
                 End If
                 Tester.Print rv1, " \\  cycle read  test"
                Call LabelMenu(1, rv1, rv0)
                
                '   CardResult = DO_WritePort(card, Channel_P1A, &HFA)
                 '  Call MsecDelay(2#)
                
                    LBA = OldLBa
                  If rv1 = 1 Then
                           rv2 = 0
                            ReaderExist = 0
                            ClosePipe
                             rv2 = CBWTest_New(0, rv1, "9360")
                          ClosePipe
                          
                            Tester.Print rv2, " \\   AU9360 test"
                   End If
                    
                Call LabelMenu(2, rv2, rv1)
                
                 If rv0 * rv1 * rv2 = PASS And TestAgainFlag = 0 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &HFE)
                   TestAgainFlag = 1
                   
                '   CardResult = DO_WritePort(card, Channel_P1CH, &H6)
                  
                 
                   GoTo SWITCH_TEST
                   
                 End If
                
               
                   
               '  Call PowerSet2(2, "5.0", "0.5", 1, "5.0", "0.5", 1)
                
                 
                 
AU6256SortigResult:
                       CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' open hub power
                   
                      If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                           AU6256Unknow = AU6256Unknow + 1
                           If AU6256Unknow > 2 Then
                             Shell "cmd /c shutdown -r  -t 0", vbHide
                           End If
                           
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
                       
                            
                        ElseIf rv1 * rv0 * rv2 = PASS Then
                             TestResult = "PASS"
                             AU6256Unknow = 0
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub

Public Sub AU6256XLS1CSortingSub()
On Error Resume Next
Dim k As Long
Dim OldTime
Dim CycleTime
Dim TestCycle As Integer
TestCycle = 2500
Dim i As Integer
Dim Vol As Single
Dim TestAgainFlag As Byte
TestAgainFlag = 0
Tester.Cls
                    If PCI7248InitFinish = 0 Then
                  PCI7248ExistAU6256
               End If
              
                   CardResult = DO_WritePort(card, Channel_P1A, &H3)
               '  Call PowerSet2(1, "5.0", "0.5", 1, "5.0", "0.5", 1)
                  CardResult = DO_WritePort(card, Channel_P1C, &H0)  ' open hub power
                 
SWITCH_TEST:
                 rv0 = 0
                 rv1 = 0
                 ReaderExist = 0
                 If TestAgainFlag = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H4)
                  Else
                  Call MsecDelay(0.8)
                  End If
              
             
                
                 Call MsecDelay(1.5)
                
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
                 Call MsecDelay(2#)
                 
                  HubPort = 1
                 ClosePipe
                    rv0 = CBWTest_New_AU6256X2_1Sorting(0, 1, "1307")
                 ClosePipe
                
                 
                 
                  Tester.Print rv0, " \\initial test"
                  Call LabelMenu(0, rv0, 1)
             ' Sorting Test
             
                 If rv0 = 1 And TestAgainFlag = 1 Then
                 rv1 = 1
                    Else
                
                 rv1 = 4
               End If
                 
                If rv0 = 1 And TestAgainFlag = 0 Then
                    OldLBa = LBA
                    
                 
                    For i = 1 To TestCycle
                        OldTime = Timer
                        OpenPipe
                        LBA = LBA + 1
                            
                            rv1 = Read_DataAU6430Sorting(LBA, 0, 512)
                        ClosePipe
                        CycleTime = Timer - OldTimer
                        If rv1 <> 1 And Abs(CycleTime) > 3 Then
                           Tester.Print "fail Clycle"; i
                        
                            Exit For
                        End If
                    Next i
                    If rv1 <> 1 Then
                       rv1 = 2
                         
                    End If
      
                End If
                 If rv1 <> 1 Then
                       rv1 = 2
                        
                 End If
                 Tester.Print rv1, " \\  cycle read  test"
                Call LabelMenu(1, rv1, rv0)
                    LBA = OldLBa
                  If rv1 = 1 Then
                           rv2 = 0
                            ReaderExist = 0
                            ClosePipe
                             rv2 = CBWTest_New(0, rv1, "9360")
                          ClosePipe
                          
                            Tester.Print rv2, " \\   AU9360 test"
                   End If
                    
                Call LabelMenu(2, rv2, rv1)
                
                 If rv0 * rv1 * rv2 = PASS And TestAgainFlag = 0 Then
                   TestAgainFlag = 1
                   
                   CardResult = DO_WritePort(card, Channel_P1A, &H6)
                  
                 
                   GoTo SWITCH_TEST
                 End If
                
               
                    
                   
              '   Call PowerSet2(2, "5.0", "0.5", 1, "5.0", "0.5", 1)
                
                 
                 
AU6371DLResult:

                    CardResult = DO_WritePort(card, Channel_P1C, &HFF)  ' open hub power
                     
                      If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                              AU6256Unknow = AU6256Unknow + 1
                           If AU6256Unknow > 2 Then
                             Shell "cmd /c shutdown -r  -t 0", vbHide
                           End If
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
                       
                            
                        ElseIf rv1 * rv0 * rv2 = PASS Then
                             TestResult = "PASS"
                              AU6256Unknow = 0
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub


Public Sub AU6256XLS1ASortingSub()
On Error Resume Next
Dim k As Long
Dim OldTime
Dim CycleTime
Dim TestCycle As Integer
TestCycle = 2500
Dim i As Integer
Dim Vol As Single
Dim TestAgainFlag As Byte
TestAgainFlag = 0
Tester.Cls
                   CardResult = DO_WritePort(card, Channel_P1A, &H3)
                 'Call PowerSet2(1, "5.0", "0.5", 1, "5.0", "0.5", 1)
SWITCH_TEST:
                 rv0 = 0
                 rv1 = 0
                 ReaderExist = 0
                 If TestAgainFlag = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H4)
                  Else
                  Call MsecDelay(0.8)
                  End If
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
              
             
                
                 Call MsecDelay(1.5)
                
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
                 Call MsecDelay(2#)
                 
                  HubPort = 1
                 ClosePipe
                    rv0 = CBWTest_New_AU6256X2_1Sorting(0, 1, "1307")
                 ClosePipe
                
                 
                 
                  Tester.Print rv0, " \\initial test"
                  Call LabelMenu(0, rv0, 1)
             ' Sorting Test
             
                 If rv0 = 1 And TestAgainFlag = 1 Then
                 rv1 = 1
                    Else
                
                 rv1 = 4
               End If
                 
                If rv0 = 1 And TestAgainFlag = 0 Then
                    OldLBa = LBA
                    
                 
                    For i = 1 To TestCycle
                        OldTime = Timer
                        OpenPipe
                        LBA = LBA + 1
                            
                            rv1 = Read_DataAU6430Sorting(LBA, 0, 512)
                        ClosePipe
                        CycleTime = Timer - OldTimer
                        If rv1 <> 1 And Abs(CycleTime) > 3 Then
                           Tester.Print "fail Clycle"; i
                        
                            Exit For
                        End If
                    Next i
                    If rv1 <> 1 Then
                       rv1 = 2
                         
                    End If
                    
                     
                    
                   
                End If
                
                
               
              
                 If rv1 <> 1 Then
                       rv1 = 2
                        
                 End If
                 Tester.Print rv1, " \\  cycle read  test"
                Call LabelMenu(1, rv1, rv0)
                    LBA = OldLBa
                  If rv1 = 1 Then
                           rv2 = 0
                            ReaderExist = 0
                            ClosePipe
                             rv2 = CBWTest_New(0, rv1, "9360")
                          ClosePipe
                          
                            Tester.Print rv2, " \\   AU9360 test"
                   End If
                    
                Call LabelMenu(2, rv2, rv1)
                
                 If rv0 * rv1 * rv2 = PASS And TestAgainFlag = 0 Then
                   TestAgainFlag = 1
                   
                   CardResult = DO_WritePort(card, Channel_P1A, &H6)
                  
                 
                   GoTo SWITCH_TEST
                 End If
                
               
                    
                   
                 'Call PowerSet2(2, "5.0", "0.5", 1, "5.0", "0.5", 1)
                
                 
                 
AU6371DLResult:
                    Call MsecDelay(0.3)
                      
                      If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                           AU6256Unknow = AU6256Unknow + 1
                            If AU6256Unknow > 2 Then
                            ' Shell "cmd /c shutdown -r  -t 0", vbHide
                           End If
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
                       
                            
                        ElseIf rv1 * rv0 * rv2 = PASS Then
                             TestResult = "PASS"
                              AU6256Unknow = 0
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub

Public Sub AU6256XLS10SortingSub_old()

Dim k As Long
Dim OldTime
Dim CycleTime
Dim TestCycle As Integer
TestCycle = 2500
Dim i As Integer
Dim Vol As Single
Tester.Cls
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
              
   
                
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
                 Call MsecDelay(2#)
                
                 ClosePipe
                    rv0 = CBWTest_New_AU6430Sorting(0, 1, "1307")
                 ClosePipe
                 
                  Tester.Print rv0, " \\initial test"
                  Call LabelMenu(0, rv0, 1)
             ' Sorting Test
                 
                If rv0 = 1 Then
                    OldLBa = LBA
                    
                 
                    For i = 1 To TestCycle
                        OldTime = Timer
                        OpenPipe
                        LBA = LBA + 1
                            
                            rv1 = Read_DataAU6430Sorting(LBA, 0, 512)
                        ClosePipe
                        CycleTime = Timer - OldTimer
                        If rv1 <> 1 And Abs(CycleTime) > 3 Then
                           Tester.Print "fail Clycle"; i
                        
                            Exit For
                        End If
                    Next i
                    If rv1 <> 1 Then
                       rv1 = 2
                         
                    End If
                    
                     
                    
                   
                End If
                
                
                 If rv1 <> 1 Then
                       rv1 = 2
                        
                 End If
                 Tester.Print rv1, " \\ 3.6V ~ 3.31 V cycle read  test"
                Call LabelMenu(1, rv1, rv0)
                  
      
            
              
               LBA = OldLBa
                 
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
                       
                            
                        ElseIf rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub
Public Sub AU6256XLS10SortingSub()

'========================================================================
    '2010/5/6
    'Skip GPIB Power-control
'========================================================================
Dim k As Long
Dim OldTime
Dim CycleTime
Dim TestCycle As Integer
TestCycle = 2500
Dim i As Integer
Dim Vol As Single
Dim UsbSpeedTestResult As Integer
Dim XLS14_Flag As Integer
Tester.Cls

UsbSpeedTestResult = 0
XLS14_Flag = 0




                'Call PowerSet2(2, "5.0", "0.5", 1, "5.0", "0.5", 1)
                'Call MsecDelay(0.3)
                'Call PowerSet2(1, "5.0", "0.5", 1, "5.0", "0.5", 1)
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
SPEED_RT:
                

              
                CardResult = DO_WritePort(card, Channel_P1A, &HFE)
                fnScsi2usb2K_KillEXE
                MsecDelay (0.3)
                
                CardResult = DO_WritePort(card, Channel_P1A, &HFC)
                MsecDelay (0.5)
                
                 
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
                 Call MsecDelay(2#)
                
                 ClosePipe

                    rv0 = CBWTest_New_AU6430Sorting(0, 1, "1307")
                 ClosePipe
                 
                    If UsbSpeedTestResult <> 0 And XLS14_Flag = 0 Then      'RT speend error when first time fail
                       GoTo SPEED_RT
                       XLS21_Flag = 1
                    End If
                 
                  Tester.Print rv0, " \\initial test"
                  Call LabelMenu(0, rv0, 1)
             ' Sorting Test
                 
                If rv0 = 1 Then
                    OldLBa = LBA
                    
                    
                    For i = 1 To TestCycle
                        OldTime = Timer
                        OpenPipe
                        LBA = LBA + 1
                            
                            rv1 = Read_DataAU6430Sorting(LBA, 0, 512)
                        ClosePipe
                        CycleTime = Timer - OldTimer
                        If rv1 <> 1 And Abs(CycleTime) > 3 Then
                           Tester.Print "2nd fail Clycle"; i
                            Exit For
                        End If
                    Next i
                    
                    
                    If rv1 <> 1 Then
                       rv1 = 2
                         
                    End If
                    
                End If

                 If rv1 <> 1 Then
                       rv1 = 2
                        
                 End If
                 Tester.Print rv1, " \\ 3.6V ~ 3.31 V cycle read  test"
                Call LabelMenu(1, rv1, rv0)
                
                
           '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                MsecDelay (0.3)
      
            
              
               LBA = OldLBa
                 
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
                       
                            
                        ElseIf rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub
Public Sub AU6256XLS17SortingSub()

Dim k As Long
Dim OldTime
Dim CycleTime
Dim TestCycle As Integer
TestCycle = 2500
Dim i As Integer
Dim Vol As Single
Tester.Cls
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
              
   
                
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
                 Call MsecDelay(2#)
                
                 ClosePipe
                    rv0 = CBWTest_New_AU6256XLS17(0, 1, "6335")
                 ClosePipe
                 
                  Tester.Print rv0, " \\initial test"
                  Call LabelMenu(0, rv0, 1)
             ' Sorting Test
                 
               
                
                
                
                
      
            
              
               LBA = OldLBa
                 
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
                       
                            
                        ElseIf rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub

Public Sub AU6256XLS16SortingSub()

Dim k As Long
Dim OldTime
Dim CycleTime
Dim TestCycle As Integer
TestCycle = 2500
Dim i As Integer
Dim Vol As Single
Tester.Cls
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
              
   
                
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
                 Call MsecDelay(2#)
                
                 ClosePipe
                   ' rv0 = CBWTest_New_AU6430Sorting(0, 1, "1307")
                       rv0 = CBWTest_New_AU6256XLS17(0, 1, "1307")
                 ClosePipe
                 
                  Tester.Print rv0, " \\initial test"
                  Call LabelMenu(0, rv0, 1)
             ' Sorting Test
                 
               
                
                
                
                
      
            
              
               LBA = OldLBa
                 
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
                       
                            
                        ElseIf rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub
Public Sub AU6256XLS13SortingSub()

Dim k As Long
Dim OldTime
Dim CycleTime
Dim TestCycle As Integer
TestCycle = 2500
Dim i As Integer
Dim Vol As Single
Tester.Cls
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
              
   
                
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
                 Call MsecDelay(2#)
                HubPort = 1
                 ClosePipe
                    rv0 = CBWTest_New_AU6256X2Sorting(0, 1, "1307")
                 ClosePipe
                 
                  Tester.Print rv0, " \\initial test"
                  Call LabelMenu(0, rv0, 1)
             ' Sorting Test
                 
                If rv0 = 1 Then
                    OldLBa = LBA
                    
                 
                    For i = 1 To TestCycle
                        OldTime = Timer
                        OpenPipe
                        LBA = LBA + 1
                            
                            rv1 = Read_DataAU6430Sorting(LBA, 0, 512)
                        ClosePipe
                        CycleTime = Timer - OldTimer
                        If rv1 <> 1 And Abs(CycleTime) > 3 Then
                           Tester.Print "fail Clycle"; i
                        
                            Exit For
                        End If
                    Next i
                    If rv1 <> 1 Then
                       rv1 = 2
                         
                    End If
                    
                     
                    
                   
                End If
                
                
                 If rv1 <> 1 Then
                       rv1 = 2
                        
                 End If
                 Tester.Print rv1, " \\   cycle read  test"
                Call LabelMenu(1, rv1, rv0)
                  
              
               
              
               LBA = OldLBa
               
                  rv2 = 0
                   ClosePipe
                    rv2 = CBWTest_New(0, rv1, "6335")
                 ClosePipe
                   Tester.Print rv2, " \\   AU6335 test"
                Call LabelMenu(2, rv2, rv1)
                  
               
               
               
                 
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
                       
                            
                        ElseIf rv1 * rv0 * rv2 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub

Public Sub AU6256XLS12SortingSub()

Dim k As Long
Dim OldTime
Dim CycleTime
Dim TestCycle As Integer
TestCycle = 2500
Dim i As Integer
Dim Vol As Single
Tester.Cls
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
              
   
                
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
                 Call MsecDelay(2#)
                HubPort = 1
                 ClosePipe
                    rv0 = CBWTest_New_AU6256X2Sorting(0, 1, "1307")
                 ClosePipe
                 
                  Tester.Print rv0, " \\initial test"
                  Call LabelMenu(0, rv0, 1)
             ' Sorting Test
                 
                If rv0 = 1 Then
                    OldLBa = LBA
                    
                 
                    For i = 1 To TestCycle
                        OldTime = Timer
                        OpenPipe
                        LBA = LBA + 1
                            
                            rv1 = Read_DataAU6430Sorting(LBA, 0, 512)
                        ClosePipe
                        CycleTime = Timer - OldTimer
                        If rv1 <> 1 And Abs(CycleTime) > 3 Then
                           Tester.Print "fail Clycle"; i
                        
                            Exit For
                        End If
                    Next i
                    If rv1 <> 1 Then
                       rv1 = 2
                         
                    End If
                    
                     
                    
                   
                End If
                
                
                 If rv1 <> 1 Then
                       rv1 = 2
                        
                 End If
                 Tester.Print rv1, " \\   cycle read  test"
                Call LabelMenu(1, rv1, rv0)
                  
              
               
              
               LBA = OldLBa
               
                  rv2 = 0
                   ClosePipe
                    rv2 = CBWTest_New(0, rv1, "9360")
                 ClosePipe
                   Tester.Print rv2, " \\   AU9360 test"
                Call LabelMenu(2, rv2, rv1)
                  
               
               
               
                 
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
                       
                            
                        ElseIf rv1 * rv0 * rv2 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub
Public Sub AU6256XLS11SortingSub()

Dim k As Long
Dim OldTime
Dim CycleTime
Dim TestCycle As Integer
TestCycle = 5000
Dim i As Integer
Dim Vol As Single
Tester.Cls
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
              
   
                
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
                 Call MsecDelay(2#)
                
                 ClosePipe
                    rv0 = CBWTest_New_AU6430Sorting(0, 1, "1307")
                 ClosePipe
                 
                  Tester.Print rv0, " \\initial test"
                  Call LabelMenu(0, rv0, 1)
             ' Sorting Test
                 
                If rv0 = 1 Then
                    OldLBa = LBA
                    
                 
                    For i = 1 To TestCycle
                        OldTime = Timer
                        OpenPipe
                        LBA = LBA + 1
                            
                            rv1 = Read_DataAU6430Sorting(LBA, 0, 512)
                        ClosePipe
                        CycleTime = Timer - OldTimer
                        If rv1 <> 1 And Abs(CycleTime) > 3 Then
                           Tester.Print "fail Clycle"; i
                        
                            Exit For
                        End If
                    Next i
                    If rv1 <> 1 Then
                       rv1 = 2
                         
                    End If
                    
                     
                    
                   
                End If
                
                
                 If rv1 <> 1 Then
                       rv1 = 2
                        
                 End If
                 Tester.Print rv1, " \\   cycle read  test"
                Call LabelMenu(1, rv1, rv0)
                  
      
            
              
               LBA = OldLBa
                 
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
                       
                            
                        ElseIf rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub

Public Sub AU6433EFS10SortingSub()

Dim k As Long
Dim OldTime
Dim CycleTime
Dim TestCycle As Integer
TestCycle = 100
Dim i As Integer
Dim Vol As Single
Tester.Cls
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
              
   
                
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
                 PowerSet (1)
                 
                   CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                  Call MsecDelay(AU6371EL_BootTime * 2 + 1.2)
             
              'Initial test
              
                 ClosePipe
                    rv0 = CBWTest_New_AU6430Sorting(0, 1, "058f")
                 ClosePipe
                 
                  Tester.Print rv0, " \\initial test"
                  Call LabelMenu(0, rv0, 1)
             ' Sorting Test
                 
                If rv0 = 1 Then
                    OldLBa = LBA
                    For Vol = 3.6 To 3.305 Step -0.01
                    
                    Call GPIBWrite("VSET1 " & CStr(Vol))
                    DoEvents
                    For i = 1 To TestCycle
                        OldTime = Timer
                        OpenPipe
                        LBA = LBA + 1
                            
                            rv1 = Read_DataAU6430Sorting(LBA, 0, 512)
                        ClosePipe
                        CycleTime = Timer - OldTimer
                        If rv1 <> 1 And Abs(CycleTime) > 3 Then
                           Tester.Print "fail Clycle"; i
                        
                            Exit For
                        End If
                    Next i
                    Tester.Print "Vol="; Vol; " RV="; rv1
                    If rv1 <> 1 Then
                       rv1 = 2
                        Exit For
                    End If
                    Next Vol
                     
                    
                   
                End If
                
                
                 If rv1 <> 1 Then
                       rv1 = 2
                        
                 End If
                 Tester.Print rv1, " \\ 3.6V ~ 3.31 V cycle read  test"
                Call LabelMenu(1, rv1, rv0)
                  
                
                  If rv1 = 1 Then
                  Tester.Cls
                    OldLBa = LBA
                    For Vol = 3.3 To 2.995 Step -0.01
                    
                    Call GPIBWrite("VSET1 " & CStr(Vol))
                    DoEvents
                    For i = 1 To TestCycle
                        OldTime = Timer
                        OpenPipe
                        LBA = LBA + 1
                            
                            rv2 = Read_DataAU6430Sorting(LBA, 0, 512)
                        ClosePipe
                        CycleTime = Timer - OldTimer
                        If rv2 <> 1 And Abs(CycleTime) > 3 Then
                           Tester.Print "fail Clycle"; i
                        
                            Exit For
                        End If
                    Next i
                    Tester.Print "Vol2="; Vol; " RV2="; rv2
                    If rv2 <> 1 Then
                       rv2 = 2
                        Exit For
                    End If
                    Next Vol
                     
                    
                   
                End If
                
                  If rv2 <> 1 Then
                       rv2 = 2
                        
                 End If
                 Tester.Print rv2, " \\3.3 V ~ 3.0 V cycle read  test"
                Call LabelMenu(2, rv2, rv1)
              ' test speed
                
              
              
              
              
              If rv2 = 1 Then
                Call MsecDelay(3)
                      ReaderExist = 0
                ClosePipe
                  rv3 = CBWTest_New_AU6430Sorting(0, 1, "058f")
                If i > 0 Then
                    If rv3 <> 0 Then
                        rv3 = TestUnitSpeed(0)
        
                        If rv3 = 0 Then
                           
                           rv3 = 2
                           UsbSpeedTestResult = 2
                          
                        End If
                    End If
                    
                 Else
                    If i <> TestCycle + 1 Then
                      rv3 = 2
                    End If
               End If
               
               If rv3 <> 1 Then
                rv3 = 2
               End If
                 
               
                
                ClosePipe
                Call LabelMenu(3, rv3, rv2)
                Tester.Print rv3, " \\Final Speed check"
              End If
              
               LBA = OldLBa
                  CardResult = DO_WritePort(card, Channel_P1A, &H80)
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
                       
                            
                        ElseIf rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub
Public Sub AU6430QLS11SortingSub()

Dim k As Long
Dim OldTime
Dim CycleTime
Dim TestCycle As Integer
TestCycle = 100
Dim i As Integer
Dim Vol As Single
Tester.Cls
                
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
              
                 
       
                 PowerSet (1316)
                 
                
                  Call MsecDelay(AU6371EL_BootTime * 2 + 1.2)
             
              'Initial test
              
                 ClosePipe
                    rv0 = CBWTest_New_AU6430Sorting(0, 1, "058f")
                 ClosePipe
                 
                  Tester.Print rv0, " \\initial test"
                  Call LabelMenu(0, rv0, 1)
             ' Sorting Test
                 
                If rv0 = 1 Then
                    OldLBa = LBA
                    For Vol = 3.16 To 3.125 Step -0.01
                    
                    Call GPIBWrite("VSET1 " & CStr(Vol))
                    DoEvents
                    For i = 1 To TestCycle
                        OldTime = Timer
                        OpenPipe
                        LBA = LBA + 1
                            
                            rv1 = Read_DataAU6430Sorting(LBA, 0, 512)
                        ClosePipe
                        CycleTime = Timer - OldTimer
                         
                        If rv1 <> 1 And Abs(CycleTime) > 3 Then
                           Tester.Print "fail Clycle"; i
                        
                            Exit For
                        End If
                    Next i
                    Tester.Print "Vol="; Vol; " RV="; rv1
                    If rv1 <> 1 Then
                       rv1 = 2
                        Exit For
                    End If
                    Next Vol
                     
                    
                   
                End If
                
                
                 If rv1 <> 1 Then
                       rv1 = 2
                        
                 End If
                 Tester.Print rv1, " \\ 3.16V ~ 3.13 V cycle read  test"
                Call LabelMenu(1, rv1, rv0)
              ' test speed
                
              
               If rv1 = 1 Then
                    OldLBa = LBA
                    For Vol = 3.12 To 3.095 Step -0.01
                    
                    Call GPIBWrite("VSET1 " & CStr(Vol))
                    DoEvents
                    For i = 1 To TestCycle
                        OldTime = Timer
                        OpenPipe
                        LBA = LBA + 1
                            
                            rv2 = Read_DataAU6430Sorting(LBA, 0, 512)
                        ClosePipe
                        CycleTime = Timer - OldTimer
                        If rv2 <> 1 And Abs(CycleTime) > 3 Then
                           Tester.Print "fail Clycle"; i
                        
                            Exit For
                        End If
                    Next i
                    Tester.Print "Vol2="; Vol; " RV2="; rv1
                    If rv2 <> 1 Then
                       rv2 = 2
                        Exit For
                    End If
                    Next Vol
                     
                    
                   
                End If
                
                  If rv2 <> 1 Then
                       rv2 = 2
                        
                 End If
                 Tester.Print rv2, " \\3.12 V ~ 3.10V cycle read  test"
                Call LabelMenu(2, rv2, rv1)
              
              
              
              
              
              If rv2 = 1 Then
                Call MsecDelay(3)
                      ReaderExist = 0
                ClosePipe
                  rv3 = CBWTest_New_AU6430Sorting(0, 1, "058f")
                If i > 0 Then
                    If rv3 <> 0 Then
                        rv3 = TestUnitSpeed(0)
        
                        If rv3 = 0 Then
                           
                           rv3 = 2
                           UsbSpeedTestResult = 2
                          
                        End If
                    End If
                    
                 Else
                    If i <> TestCycle + 1 Then
                      rv3 = 2
                    End If
               End If
               
               If rv3 <> 1 Then
                rv3 = 2
               End If
                 
               
                
                ClosePipe
                Call LabelMenu(3, rv3, rv2)
                Tester.Print rv3, " \\Final Speed check"
              End If
              
               LBA = OldLBa
                
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
                       
                            
                        ElseIf rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub

Public Sub AU6256XLS20SortingSub()

'========================================================================
    'copy from AU6256XLS10Sorting
    'read 2500 cycle => P1 UFD PWR OFF => P1 ON => 2500 cycle
    'Purpose to sorting Pasonic case
'========================================================================
Dim k As Long
Dim OldTime
Dim CycleTime
Dim TestCycle As Integer
TestCycle = 2500
Dim i As Integer
Dim Vol As Single
Dim UsbSpeedTestResult As Integer
Dim XLS21_Flag As Integer
Tester.Cls

UsbSpeedTestResult = 0
XLS21_Flag = 0




                Call PowerSet2(2, "5.0", "0.5", 1, "5.0", "0.5", 1)
                Call MsecDelay(0.3)
                Call PowerSet2(1, "5.0", "0.5", 1, "5.0", "0.5", 1)
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
SPEED_RT:
                

              
                CardResult = DO_WritePort(card, Channel_P1A, &HFE)
                fnScsi2usb2K_KillEXE
                MsecDelay (0.2)
                
                CardResult = DO_WritePort(card, Channel_P1A, &HFC)
                MsecDelay (0.5)
                
                'Open "D:\Record.txt" For Append As #1
                
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
                 Call MsecDelay(2#)
                
                 ClosePipe

                    rv0 = CBWTest_New_AU6430Sorting(0, 1, "1307")
                 ClosePipe
                 
                    If UsbSpeedTestResult <> 0 And XLS21_Flag = 0 Then      'RT speend error when first time fail
                       GoTo SPEED_RT
                       XLS21_Flag = 1
                    End If
                 
                  Tester.Print rv0, " \\initial test"
                  Call LabelMenu(0, rv0, 1)
             ' Sorting Test
                 
                If rv0 = 1 Then
                    OldLBa = LBA
                    
                 
                    For i = 1 To TestCycle
                        OldTime = Timer
                        OpenPipe
                        LBA = LBA + 1
                            
                            rv1 = Read_DataAU6430Sorting(LBA, 0, 512)
                        ClosePipe
                        CycleTime = Timer - OldTimer
                        If rv1 <> 1 And Abs(CycleTime) > 3 Then
                           Tester.Print "1st: fail Clycle"; i
                           'Print #1, "1ST: " & i
                           
                            Exit For
                        End If
                    Next i
                    
                    If rv1 <> 1 Then
                       rv1 = 2
                       GoTo Exit_2nd_LOOP
                    End If
                    
                Call LabelMenu(0, rv1, 1)
                
                CardResult = DO_WritePort(card, Channel_P1A, &HFE)  'close UFD
                fnScsi2usb2K_KillEXE
                MsecDelay (0.2)
                
                
                
                CardResult = DO_WritePort(card, Channel_P1A, &HFC)
                MsecDelay (0.5)
                    
                    
                    
                    For i = 1 To TestCycle
                        OldTime = Timer
                        OpenPipe
                        LBA = LBA + 1
                            
                            rv1 = Read_DataAU6430Sorting(LBA, 0, 512)
                        ClosePipe
                        CycleTime = Timer - OldTimer
                        If rv1 <> 1 And Abs(CycleTime) > 3 Then
                           Tester.Print "2nd fail Clycle"; i
                           'Print #1, "2nd: " & i
                            Exit For
                        End If
                    Next i
                    
                    
                    If rv1 <> 1 Then
                       rv1 = 2
                         
                    End If
                    
                End If
                
Exit_2nd_LOOP:

                 If rv1 <> 1 Then
                       rv1 = 2
                        
                 End If
                 Tester.Print rv1, " \\ 3.6V ~ 3.31 V cycle read  test"
                Call LabelMenu(1, rv1, rv0)
                
                
           '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                MsecDelay (0.3)
      
            
              
               LBA = OldLBa
                 
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
                       
                            
                        ElseIf rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                             'Print #1, "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
            'Close #1
End Sub


Public Sub AU6256XLS21SortingSub()

'========================================================================
    'copy from AU6256XLS10Sorting
    'read 2500 cycle => P1 UFD PWR OFF => P1 ON => 2500 cycle
    'Purpose to sorting Pasonic case
'========================================================================
Dim k As Long
Dim OldTime
Dim CycleTime
Dim TestCycle As Integer
TestCycle = 2500
Dim i As Integer
Dim Vol As Single
Dim UsbSpeedTestResult As Integer
Dim XLS21_Flag As Integer
Tester.Cls

UsbSpeedTestResult = 0
XLS21_Flag = 0




                'Call PowerSet2(2, "5.0", "0.5", 1, "5.0", "0.5", 1)
                'Call MsecDelay(0.3)
                'Call PowerSet2(1, "5.0", "0.5", 1, "5.0", "0.5", 1)
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
SPEED_RT:
                

              
                CardResult = DO_WritePort(card, Channel_P1A, &HFE)
                fnScsi2usb2K_KillEXE
                MsecDelay (0.3)
                
                CardResult = DO_WritePort(card, Channel_P1A, &HFC)
                MsecDelay (0.5)
                
                'Open "D:\Record.txt" For Append As #1
                
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
                 Call MsecDelay(2#)
                
                 ClosePipe

                    rv0 = CBWTest_New_AU6430Sorting(0, 1, "1307")
                 ClosePipe
                 
                    If UsbSpeedTestResult <> 0 And XLS21_Flag = 0 Then      'RT speend error when first time fail
                       GoTo SPEED_RT
                       XLS21_Flag = 1
                    End If
                 
                  Tester.Print rv0, " \\initial test"
                  Call LabelMenu(0, rv0, 1)
             ' Sorting Test
                 
                If rv0 = 1 Then
                    OldLBa = LBA
                    
                 
                    For i = 1 To TestCycle
                        OldTime = Timer
                        OpenPipe
                        LBA = LBA + 1
                            
                            rv1 = Read_DataAU6430Sorting(LBA, 0, 512)
                        ClosePipe
                        CycleTime = Timer - OldTimer
                        If rv1 <> 1 And Abs(CycleTime) > 3 Then
                           Tester.Print "1st: fail Clycle"; i
                           'Print #1, "1ST: " & i
                           
                            Exit For
                        End If
                    Next i
                    
                    If rv1 <> 1 Then
                       rv1 = 2
                       GoTo Exit_2nd_LOOP
                    End If
                    
                Call LabelMenu(0, rv1, 1)
                
                CardResult = DO_WritePort(card, Channel_P1A, &H0)  'close UFD
                fnScsi2usb2K_KillEXE
                MsecDelay (0.4)
                
                
                
                CardResult = DO_WritePort(card, Channel_P1A, &HFC)
                MsecDelay (0.8)
                    
                    
                    
                    For i = 1 To TestCycle
                        OldTime = Timer
                        OpenPipe
                        LBA = LBA + 1
                            
                            rv1 = Read_DataAU6430Sorting(LBA, 0, 512)
                        ClosePipe
                        CycleTime = Timer - OldTimer
                        If rv1 <> 1 And Abs(CycleTime) > 3 Then
                           Tester.Print "2nd fail Clycle"; i
                           'Print #1, "2nd: " & i
                            Exit For
                        End If
                    Next i
                    
                    
                    If rv1 <> 1 Then
                       rv1 = 2
                         
                    End If
                    
                End If
                
Exit_2nd_LOOP:

                 If rv1 <> 1 Then
                       rv1 = 2
                        
                 End If
                 Tester.Print rv1, " \\ 3.6V ~ 3.31 V cycle read  test"
                Call LabelMenu(1, rv1, rv0)
                
                
           '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                MsecDelay (0.3)
      
            
              
               LBA = OldLBa
                 
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
                       
                            
                        ElseIf rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                             'Print #1, "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
            'Close #1
End Sub

Public Sub AU6256XLS31SortingSub()

'========================================================================
    'copy from AU6256XLS21Sorting
'========================================================================
Dim k As Long
Dim OldTime
Dim CycleTime
TestCycle = 2500
Dim i As Integer
Dim UsbSpeedTestResult As Integer
Dim ReaderPID As String
Dim UFDPID As String

Tester.Cls
UsbSpeedTestResult = 0


                If PCI7248InitFinish = 0 Then
                    PCI7248Exist
                End If
             
                LBA = LBA + 1
                         
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                rv5 = 0
                rv6 = 0
                rv7 = 0
                ReaderPID = "058f"
                UFDPID = "1307"
                
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '=========================================
                '    POWER on
                '=========================================
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H0)
                Call MsecDelay(2#)
                
                Tester.Print "Begin UFD & Reader Test..."
                
                ClosePipe
                HubPort = 1
                rv0 = CBWTest_New_AU6256X2_1Sorting(0, 1, UFDPID)
                'ClosePipe
                 
                If rv0 = 1 Then
                    rv0 = Read_DataAU6430Sorting(LBA, 0, 512)
                End If
                 
                'If rv0 = 1 Then
                '    For i = 1 To TestCycle
                '        OldTime = Timer
                '        OpenPipe
                '        LBA = LBA + 1
                '        rv0 = Read_DataAU6430Sorting(LBA, 0, 512)
                '        ClosePipe
                '        If (rv0 <> 1) Or (Abs(OldTime - Timer) > 2) Then
                '            rv0 = 3
                '            Exit For
                '        End If
                '    Next i
                'End If
                
                
                If rv0 = 1 Then
                    Tester.Print "UFD Test PASS !"
                Else
                    Tester.Print "UFD Test Fail !"
                End If
                
                Call LabelMenu(0, rv0, 1)
                
                If rv0 = 1 Then
                    rv1 = CBWTest_New(0, 1, ReaderPID)
                    ClosePipe
                    
                    If rv1 = 1 Then
                        Tester.Print "Reader Test PASS !"
                    Else
                        Tester.Print "Reader Test Fail !"
                    End If
                End If
                
                Call LabelMenu(0, rv1, rv0)
                
                OldLBa = LBA
                 
AU6371DLResult:
                
                CardResult = DO_WritePort(card, Channel_P1A, &HFF)  'Close power
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
                        ElseIf rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                        End If
End Sub

Public Sub AU6256XLS32SortingSub()

'========================================================================
'  copy from AU6256XLS31Sorting
'  insert 4 ea Elecom UFD on downstream port
'  purpose to solve "321" RMA
'========================================================================
Dim k As Long
Dim OldTime
Dim CycleTime
Dim i As Integer
Dim UsbSpeedTestResult As Integer
Dim ReaderPID As String
Dim UFDPID As String

Tester.Cls
UsbSpeedTestResult = 0


                If PCI7248InitFinish = 0 Then
                    PCI7248Exist
                End If
             
                LBA = LBA + 1
                         
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                rv5 = 0
                rv6 = 0
                rv7 = 0
                UFDPID = "1307"
                
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '=========================================
                '    POWER on
                '=========================================
                
                fnScsi2usb2K_KillEXE
                CardResult = DO_WritePort(card, Channel_P1A, &H1)      'ENA ON
                Call MsecDelay(1.2)
                
                CardResult = DO_WritePort(card, Channel_P1A, &HE6)      '2 downstram port on
                Call MsecDelay(0.8)
                
                CardResult = DO_WritePort(card, Channel_P1A, &HE0)      'All downstram port on
                Call MsecDelay(0.8)
                
                '=============================== UFD1 Test Start ===============================
                
                Tester.Print "Begin UFD1 Test..."
                
                ClosePipe
                rv0 = CBWTest_New_MultiDevice(0, 1, UFDPID, 0)
                 
                If rv0 = 1 Then
                    rv0 = CBWTest_New_128_Sector_PipeReady(0, rv0)
                    ClosePipe
                End If
                 
                ClosePipe
                Tester.Print rv0, " \\UFD1 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                If rv0 = 1 Then
                    Tester.Print "UFD1 Test PASS !" & vbCrLf
                Else
                    Tester.Print "UFD1 Test Fail !" & vbCrLf
                End If
                
                Call LabelMenu(0, rv0, 1)
                
                '=============================== UFD1 Test End ===============================
                
                
                
                '=============================== UFD2 Test Start ===============================
                
                Tester.Print "Begin UFD2 Test..."
                
                ClosePipe
                rv1 = CBWTest_New_MultiDevice(0, rv0, UFDPID, 1)
                 
                If rv1 = 1 Then
                    rv1 = CBWTest_New_128_Sector_PipeReady(0, rv1)
                    ClosePipe
                End If
                 
                ClosePipe
                Tester.Print rv1, " \\UFD2 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                If rv1 = 1 Then
                    Tester.Print "UFD2 Test PASS !" & vbCrLf
                Else
                    Tester.Print "UFD2 Test Fail !" & vbCrLf
                End If
                
                Call LabelMenu(0, rv1, rv0)
                
                '=============================== UFD2 Test End ===============================
                
                
                
                '=============================== UFD3 Test Start ===============================
                
                Tester.Print "Begin UFD3 Test..."
                
                ClosePipe
                rv2 = CBWTest_New_MultiDevice(0, rv1, UFDPID, 2)
                 
                If rv2 = 1 Then
                    rv2 = CBWTest_New_128_Sector_PipeReady(0, rv2)
                    ClosePipe
                End If
                 
                ClosePipe
                Tester.Print rv2, " \\UFD3 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                If rv2 = 1 Then
                    Tester.Print "UFD2 Test PASS !" & vbCrLf
                Else
                    Tester.Print "UFD2 Test Fail !" & vbCrLf
                End If
                
                Call LabelMenu(0, rv2, rv1)
                
                '=============================== UFD3 Test End ===============================
                
                
                
                '=============================== UFD4 Test Start ===============================
                
                Tester.Print "Begin UFD4 Test..."
                
                ClosePipe
                rv3 = CBWTest_New_MultiDevice(0, rv2, UFDPID, 3)
                 
                If rv3 = 1 Then
                    rv3 = CBWTest_New_128_Sector_PipeReady(0, rv3)
                    ClosePipe
                End If
                 
                ClosePipe
                Tester.Print rv3, " \\UFD4 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                If rv3 = 1 Then
                    Tester.Print "UFD4 Test PASS !"
                Else
                    Tester.Print "UFD4 Test Fail !"
                End If
                
                Call LabelMenu(0, rv3, rv2)
                
                '=============================== UFD3 Test End ===============================
                
                OldLBa = LBA
                 
AU6371DLResult:
                
                CardResult = DO_WritePort(card, Channel_P1A, &HFF)  'Close power
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
                        ElseIf rv0 * rv1 * rv2 * rv3 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                        End If
End Sub

Public Sub AU6256XLSE2SortingSub()

'========================================================================
'  copy from AU6256XLS32SortingSub loop 10 cycle
'  insert 4 ea Elecom UFD on downstream port
'  purpose to solve "321" RMA
'========================================================================
Dim k As Long
Dim OldTime
Dim CycleTime
Dim i As Integer
Dim UsbSpeedTestResult As Integer
Dim ReaderPID As String
Dim UFDPID As String
Dim LoopCycleSet As Integer
Dim CurCycleCount As Integer

Tester.Cls
UsbSpeedTestResult = 0
LoopCycleSet = 10
CurCycleCount = 1

                If PCI7248InitFinish = 0 Then
                    PCI7248Exist
                End If
             
                LBA = LBA + 1
                         
                rv0 = 0
                rv1 = 0
                rv2 = 0
                rv3 = 0
                rv4 = 0
                rv5 = 0
                rv6 = 0
                rv7 = 0
                UFDPID = "1307"
                
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
                '=========================================
                '    POWER on
                '=========================================
                
                fnScsi2usb2K_KillEXE
                CardResult = DO_WritePort(card, Channel_P1A, &H1)      'ENA ON
                Call MsecDelay(1.2)
                
                CardResult = DO_WritePort(card, Channel_P1A, &HE6)      '2 downstram port on
                Call MsecDelay(0.8)
                
                CardResult = DO_WritePort(card, Channel_P1A, &HE0)      'All downstram port on
                Call MsecDelay(0.8)
                
                '=============================== UFD1 Test Start ===============================
                
                Tester.Print "Begin UFD1 Test..."
                
                ClosePipe
                rv0 = CBWTest_New_MultiDevice(0, 1, UFDPID, 0)
                
                Do
                    
                    If rv0 = 1 Then
                        rv0 = CBWTest_New_128_Sector_PipeReady(0, rv0)
                        If rv0 <> 1 Then
                            Tester.Print "UFD1 Fail Cycle is " & CurCycleCount
                        End If
                    End If
                
                    CurCycleCount = CurCycleCount + 1
                
                Loop Until ((rv0 <> 1) Or (CurCycleCount > LoopCycleSet))
                  
                CurCycleCount = 1
                ClosePipe
                Tester.Print rv0, " \\UFD1 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                If rv0 = 1 Then
                    Tester.Print "UFD1 Test PASS !" & vbCrLf
                Else
                    Tester.Print "UFD1 Test Fail !" & vbCrLf
                End If
                
                Call LabelMenu(0, rv0, 1)
                
                '=============================== UFD1 Test End ===============================
                
                
                
                '=============================== UFD2 Test Start ===============================
                
                Tester.Print "Begin UFD2 Test..."
                
                ClosePipe
                rv1 = CBWTest_New_MultiDevice(0, rv0, UFDPID, 1)
                 
                Do
                    
                    If rv1 = 1 Then
                        rv1 = CBWTest_New_128_Sector_PipeReady(0, rv1)
                        If rv1 <> 1 Then
                            Tester.Print "UFD2 Fail Cycle is " & CurCycleCount
                        End If
                    End If
                
                    CurCycleCount = CurCycleCount + 1
                
                Loop Until ((rv1 <> 1) Or (CurCycleCount > LoopCycleSet))
                  
                CurCycleCount = 1
                ClosePipe
                Tester.Print rv1, " \\UFD2 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                If rv1 = 1 Then
                    Tester.Print "UFD2 Test PASS !" & vbCrLf
                Else
                    Tester.Print "UFD2 Test Fail !" & vbCrLf
                End If
                
                Call LabelMenu(0, rv1, rv0)
                
                '=============================== UFD2 Test End ===============================
                
                
                
                '=============================== UFD3 Test Start ===============================
                
                Tester.Print "Begin UFD3 Test..."
                
                ClosePipe
                rv2 = CBWTest_New_MultiDevice(0, rv1, UFDPID, 2)
                 
                Do
                    
                    If rv2 = 1 Then
                        rv2 = CBWTest_New_128_Sector_PipeReady(0, rv2)
                        If rv2 <> 1 Then
                            Tester.Print "UFD3 Fail Cycle is " & CurCycleCount
                        End If
                    End If
                
                    CurCycleCount = CurCycleCount + 1
                
                Loop Until ((rv2 <> 1) Or (CurCycleCount > LoopCycleSet))
                  
                CurCycleCount = 1
                ClosePipe
                Tester.Print rv2, " \\UFD3 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                If rv2 = 1 Then
                    Tester.Print "UFD2 Test PASS !" & vbCrLf
                Else
                    Tester.Print "UFD2 Test Fail !" & vbCrLf
                End If
                
                Call LabelMenu(0, rv2, rv1)
                
                '=============================== UFD3 Test End ===============================
                
                
                
                '=============================== UFD4 Test Start ===============================
                
                Tester.Print "Begin UFD4 Test..."
                
                ClosePipe
                rv3 = CBWTest_New_MultiDevice(0, rv2, UFDPID, 3)
                 
                Do
                    
                    If rv3 = 1 Then
                        rv3 = CBWTest_New_128_Sector_PipeReady(0, rv3)
                        If rv3 <> 1 Then
                            Tester.Print "UFD4 Fail Cycle is " & CurCycleCount
                        End If
                    End If
                
                    CurCycleCount = CurCycleCount + 1
                
                Loop Until ((rv3 <> 1) Or (CurCycleCount > LoopCycleSet))
                  
                CurCycleCount = 1
                ClosePipe
                Tester.Print rv3, " \\UFD4 :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                If rv3 = 1 Then
                    Tester.Print "UFD4 Test PASS !"
                Else
                    Tester.Print "UFD4 Test Fail !"
                End If
                
                Call LabelMenu(0, rv3, rv2)
                
                '=============================== UFD3 Test End ===============================
                
                OldLBa = LBA
                 
AU6371DLResult:
                
                CardResult = DO_WritePort(card, Channel_P1A, &HFF)  'Close power
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
                        ElseIf rv0 * rv1 * rv2 * rv3 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                        End If
End Sub
