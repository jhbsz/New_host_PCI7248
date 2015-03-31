Attribute VB_Name = "AU6433MDL"
Option Explicit

Public Sub AU6433F2AMPTest()

    If ChipName = "AU6433JSF2A" Then
        Call AU6433JSF2ATestSub
    ElseIf ChipName = "AU6433JSF3A" Then
        Call AU6433JSF3ATestSub
    ElseIf ChipName = "AU6433FSF2A" Then
        Call AU6433FSF2ATestSub
    ElseIf ChipName = "AU6433FSF0A" Then
        Call AU6433FSF0ATestSub
    End If
        
End Sub
Public Sub AU6433S61MPTest()

If ChipName = "AU6433BLS2F" Then
    Call AU6433BLS2FTestSub
ElseIf ChipName = "AU6433BLS0F" Then
    Call AU6433BLS0FTestSub
End If


End Sub
Public Sub AU6433BLF24TestSub()

Call PowerSet2(1, "3.3", "0.05", 1, "3.3", "0.05", 1)
      
  Tester.Print "AU6433EF : NB mode test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
               If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
               
               ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
               '  CardResult = DO_WritePort(card, Channel_P1B, &H0)
               
                LBA = LBA + 199
                         
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
                ChipString = "vid"
                 If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
               '   Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1.3)
                     
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
                      
                      Tester.Print "rv0="; rv0
                     
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                          Tester.Print "LightON="; LightOn
                          Tester.Print "LightOFF="; LightOff
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
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
                 Call MsecDelay(0.1)
               
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
                 Call MsecDelay(0.1)
                  OpenPipe
                  rv5 = ReInitial(0)
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                If rv5 = 1 Then
                   rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                   If rv5 <> 1 Then
                      rv5 = 2
                      Tester.Print "MS bus width Fail"
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
Public Sub AU6433BLS10SortingSub()

Call PowerSet2(1, "3.3", "0.05", 1, "2.2", "0.05", 1)
      
  Tester.Print "AU6433EF : NB mode test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim CPRMMODE As Byte
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                
              
               
                '===============================================
                '  SD Card test
                '
               '   Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1.3)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
                     
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
                       
                            
                        ElseIf rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub

Public Sub AU6433BLS11SortingSub()

Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer

Call PowerSet2(1, "3.3", "0.05", 1, "1.8", "0.05", 1)
      
  Tester.Print "AU6433BL : NB mode test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim CPRMMODE As Byte
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                
              
               
                '===============================================
                '  SD Card test
                '
                 
               
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                   Call MsecDelay(1.5)
                 
                      
                    rv0 = 1
                            
                     TmpString = ""
                 
                        Do
                            DoEvents
                            Call MsecDelay(0.1)
                            TimerCounter = TimerCounter + 1
                            TmpString = GetDeviceName("058f")
                        Loop While TmpString = "" And TimerCounter < 10
                
                  
                    If TmpString = "" Then
                      rv0 = 0   ' no readerExist
                      Tester.Label9.Caption = "unknow device"
                      GoTo AU6371ELResult
                    End If
                    
                  '===========================================
                 'NO card test 2.2 V
                 '============================================
                    
                    
                    Call PowerSet2(2, "3.3", "0.05", 1, "2.2", "0.05", 1)
                    
                    Call MsecDelay(0.5)
                    Call PowerSet2(1, "3.3", "0.05", 1, "2.2", "0.05", 1)
                     Call MsecDelay(1.5)
                         Do
                            DoEvents
                            Call MsecDelay(0.1)
                            TimerCounter = TimerCounter + 1
                            TmpString = GetDeviceName("058f")
                        Loop While TmpString = "" And TimerCounter < 10
                   
                   
                    If TmpString = "" Then
                      rv0 = 2   ' no readerExist
                        Tester.Label9.Caption = "Sorting Fail"
                       
                    End If
                    
                     Tester.Label9.Caption = "PASS"
                    
                    
                    
                     
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
                       
                            
                        ElseIf rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub

Public Sub AU6433BLS13SortingSub()

Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer

Call PowerSet2(1, "3.3", "0.05", 1, "1.85", "0.05", 1)
      
  Tester.Print "AU6433BL : NB mode test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim CPRMMODE As Byte
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                
              
               
                '===============================================
                '  SD Card test
                '
                 
               
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                      
                      Call MsecDelay(1.5)
                                      
                      rv0 = 0   ' no readerExist
                      ReaderExist = 0
                     rv0 = CBWTest_New(0, 1, ChipString)
                     Tester.Print "1.85V rv0:"; rv0
                       If rv0 <> 1 Then
                        Tester.Label9.Caption = "1.8 V Sorting Fail"
                        GoTo AU6371ELResult
                    End If
                    
                  '===========================================
                 'NO card test 2.2 V
                 '============================================
                    
                    
                    Call PowerSet2(2, "3.3", "0.05", 1, "2.2", "0.05", 1)
                    
                    Call MsecDelay(0.5)
                    Call PowerSet2(1, "3.3", "0.05", 1, "2.2", "0.05", 1)
                     Call MsecDelay(1.2)
                         
                         
                     ClosePipe
                     rv0 = 0   ' no readerExist
                     ReaderExist = 0
                     rv0 = CBWTest_New(0, 1, ChipString)
                     Tester.Print "2.2V rv0:"; rv0
                     If rv0 <> 1 Then
                        Tester.Label9.Caption = "2.2 V Sorting Fail"
                        GoTo AU6371ELResult
                    End If
                    
                     Tester.Label9.Caption = "PASS"
                    
                    
                    
                     
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
                       
                            
                        ElseIf rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub

Public Sub AU6433BLS12SortingSub()

Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer

Call PowerSet2(1, "3.3", "0.05", 1, "1.80", "0.05", 1)
      
  Tester.Print "AU6433BL : NB mode test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim CPRMMODE As Byte
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                
              
               
                '===============================================
                '  SD Card test
                '
                 
               
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                      
                      Call MsecDelay(1.2)
                                      
                      rv0 = 0   ' no readerExist
                      ReaderExist = 0
                     rv0 = CBWTest_New(0, 1, ChipString)
                     Tester.Print "1.8V rv0:"; rv0
                       If rv0 <> 1 Then
                        If rv0 = 0 Then
                           Tester.Label9.Caption = "1.8 V unknow Fail"
                        Else
                           Tester.Label9.Caption = "1.8 V R/W  Fail"
                        End If
                        GoTo AU6371ELResult
                    End If
                    
                  '===========================================
                 'NO card test 2.2 V
                 '============================================
                    
                    
                    Call PowerSet2(2, "3.3", "0.05", 1, "2.2", "0.05", 1)
                    
                    Call MsecDelay(0.5)
                    Call PowerSet2(1, "3.3", "0.05", 1, "2.2", "0.05", 1)
                     Call MsecDelay(1.5)
                         
                         
                     ClosePipe
                     rv1 = 0   ' no readerExist
                     ReaderExist = 0
                     rv1 = CBWTest_New(0, 1, ChipString)
                     Tester.Print "2.2 V rv1:"; rv1
                     If rv1 <> 1 Then
                        If rv1 = 0 Then
                           Tester.Label9.Caption = "2.2 V unknow Fail"
                        Else
                           Tester.Label9.Caption = "2.2 V R/W  Fail"
                        End If
                        
                        GoTo AU6371ELResult
                    End If
                    
                     Tester.Label9.Caption = "PASS"
                    
                    
                    
                     
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
                            
                        ElseIf rv1 = UNKNOW Then
                        
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                            
                        ElseIf rv1 = WRITE_FAIL Then
                             MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv1 = READ_FAIL Then
                             MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                      
                       
                            
                        ElseIf rv0 * rv1 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
                        
                        Call PowerSet2(1, "3.3", "0.05", 1, "1.80", "0.05", 1)
End Sub
Public Sub AU6433BLF26TestSub()

Call PowerSet2(1, "3.3", "0.06", 1, "1.9", "0.06", 1)
      
  Tester.Print "AU6433EF : NB mode test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim CPRMMODE As Byte
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
              
                 
                 '================================================
                  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H3E) '0100 0000 control 2.2 V input
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(0.8)
                     
                 
                
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      
                      
                        If rv0 = 1 Then
                      
                             CardResult = DO_WritePort(card, Channel_P1A, &H80)  ' power off chip
                             Call PowerSet2(1, "3.3", "0.06", 1, "2.2", "0.06", 1) 'test 2.2 v
                             Call MsecDelay(0.4)
                        
                             CardResult = DO_WritePort(card, Channel_P1A, &H3E)
                            
                             Call MsecDelay(0.8)
                             rv0 = 0
                             ReaderExist = 0
                             rv0 = CBWTest_New(0, 1, ChipString)
                     
                      
                      Else
                                rv0 = 5   ' 1.9 V  sorting fail
                                GoTo AU6371ELResult
                      
                      End If
                      
                      
                      If rv0 = 1 Then
                            
                             CardResult = DO_WritePort(card, Channel_P1A, &H80)  ' power off chip
                             Call MsecDelay(0.4)
                             CardResult = DO_WritePort(card, Channel_P1A, &H7E)  'switch to 1.8 regulator
                             Call MsecDelay(0.8)
                             rv0 = 0
                             ReaderExist = 0
                             rv0 = CBWTest_New(0, 1, ChipString)
                       Else
                                rv0 = 6   ' 2.2 V  sorting fail
                                GoTo AU6371ELResult

                      End If
                      
                        
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                      If rv0 = 1 Then
                           rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                           If rv0 <> 1 Then
                           rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                       End If
                      
                       If rv0 <> 1 Then ' for E55 command
                          rv0 = Read_SD_SpeedE55(0, 0, 64, "8Bits")
                          If rv0 <> 1 Then
                          rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                          If rv0 = 1 Then
                            CPRMMODE = 1
                          End If
                      End If
                      ClosePipe
                      
                      Tester.Print "rv0="; rv0
                      Call MsecDelay(0.2)
                       CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                        If rv0 = 1 Then
                          If LightOn <> &HBF Then
                          Tester.Print "LightON="; LightOn
                          Tester.Print "LightOFF="; LightOff
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
              
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
                 Call MsecDelay(0.1)
               
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
                 Call MsecDelay(0.1)
                  OpenPipe
                  rv5 = ReInitial(0)
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                
                If CPRMMODE = 0 Then  ' for E54 Before
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                 Else             ' for AU6433E55 after
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476E55(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                End If
                
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                
                '=========================================================
                  
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(0.4)    'power on time
                ChipString = "vid"
                 If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                  CardResult = DO_WritePort(card, Channel_P1A, &H3E)   ' Close power
                
AU6371ELResult:
                      If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                          ElseIf rv0 = 5 Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                            
                          ElseIf rv0 = 6 Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
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
                           SDReadFail = SDReadFail + 1
                            TestResult = "SD_RF"
                        ElseIf rv2 = READ_FAIL Or rv3 = READ_FAIL Then
                           SDReadFail = SDReadFail + 1
                            TestResult = "SD_RF"
                         ElseIf rv4 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                            SDReadFail = SDReadFail + 1
                            TestResult = "SD_RF"
                        ElseIf rv4 = READ_FAIL Or rv5 = READ_FAIL Then
                           SDReadFail = SDReadFail + 1
                            TestResult = "SD_RF"
                       
                            
                        ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub

Public Sub AU6433BLF28TestSub()

Call PowerSet2(1, "3.3", "0.05", 1, "2.1", "0.05", 1)
      
  Tester.Print "AU6433EF : NB mode test , change 1.8V to 2.1V"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim CPRMMODE As Byte
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
                ChipString = "vid"
                 If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
               '   Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1.3)
                     
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
                      
                       If rv0 <> 1 Then ' for E55 command
                          rv0 = Read_SD_SpeedE55(0, 0, 64, "8Bits")
                          If rv0 <> 1 Then
                          rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                          If rv0 = 1 Then
                            CPRMMODE = 1
                          End If
                      End If
                      ClosePipe
                      
                      Tester.Print "rv0="; rv0
                     
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                          Tester.Print "LightON="; LightOn
                          Tester.Print "LightOFF="; LightOff
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
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
                 Call MsecDelay(0.1)
               
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
                 Call MsecDelay(0.1)
                  OpenPipe
                  rv5 = ReInitial(0)
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                
                If CPRMMODE = 0 Then  ' for E54 Before
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                 Else             ' for AU6433E55 after
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476E55(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
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
Public Sub AU6433BLF29TestSub()

Dim TmpLBA As Long
Dim i As Integer
Call PowerSet2(1, "3.3", "0.05", 1, "2.1", "0.05", 1)
      
  Tester.Print "AU6433EF : NB mode test , change 1.8V to 2.1V"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim CPRMMODE As Byte
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
                ChipString = "vid"
                 If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
               '   Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1.3)
                     
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
                      
                       If rv0 <> 1 Then ' for E55 command
                          rv0 = Read_SD_SpeedE55(0, 0, 64, "8Bits")
                          If rv0 <> 1 Then
                          rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                          If rv0 = 1 Then
                            CPRMMODE = 1
                          End If
                      End If
                      ClosePipe
                      
                      Tester.Print "rv0="; rv0
                     
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                          Tester.Print "LightON="; LightOn
                          Tester.Print "LightOFF="; LightOff
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                    
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      
                     TmpLBA = LBA
                     'LBA = 99
                         'For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6371ELResult
                             End If
                         'Next
                      LBA = TmpLBA
                      
'=======================================================================================
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
                 Call MsecDelay(0.1)
               
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
                 Call MsecDelay(0.1)
                  OpenPipe
                  rv5 = ReInitial(0)
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                
                If CPRMMODE = 0 Then  ' for E54 Before
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                 Else             ' for AU6433E55 after
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476E55(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
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
Public Sub AU6433BLF2ATestSub()

Dim TmpLBA As Long
Dim i As Integer

               If PCI7248InitFinish = 0 Then
                  PCI7248ExistAU6254
                  Call SetTimer_1ms
               End If

Call PowerSet2(1, "3.3", "0.05", 1, "2.1", "0.05", 1)

OS_Result = 0
rv0 = 0

CardResult = DO_WritePort(card, Channel_P1C, &H0)
                 
MsecDelay (0.3)

OpenShortTest_Result

If OS_Result <> 1 Then
    rv0 = 0                 'OS Fail
    GoTo AU6371ELResult
End If

CardResult = DO_WritePort(card, Channel_P1C, &HFF)
                       
      
  Tester.Print "Begin AU6433EF FT2 Test: NB mode test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim CPRMMODE As Byte
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
               
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
                ChipString = "vid"
                 If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
               '   Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1.3)
                     
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
                      
                       If rv0 <> 1 Then ' for E55 command
                          rv0 = Read_SD_SpeedE55(0, 0, 64, "8Bits")
                          If rv0 <> 1 Then
                          rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                          If rv0 = 1 Then
                            CPRMMODE = 1
                          End If
                      End If
                      ClosePipe
                      
                      Tester.Print "rv0="; rv0
                     
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                          Tester.Print "LightON="; LightOn
                          Tester.Print "LightOFF="; LightOff
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                    
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      If rv0 = 1 Then
                        TmpLBA = LBA
                        'LBA = 99
                         'For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6371ELResult
                             End If
                         'Next
                        LBA = TmpLBA
                      End If
'=======================================================================================
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
                 Call MsecDelay(0.1)
               
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
                 Call MsecDelay(0.1)
                  OpenPipe
                  rv5 = ReInitial(0)
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                
                If CPRMMODE = 0 Then  ' for E54 Before
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                 Else             ' for AU6433E55 after
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476E55(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
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
Public Sub AU6433BLF2CTestSub()


'201012/30 This code just for V2 S/B
'reduce NBMD test time

Dim TmpLBA As Long
Dim i As Integer

               If PCI7248InitFinish = 0 Then
                  PCI7248ExistAU6254
                  Call SetTimer_1ms
               End If
If Dir("D:\LABPC.PC") = "LABPC.PC" Then
    Call PowerSet2(1, "3.3", "0.05", 1, "2.2", "0.05", 1)
Else
    Call PowerSet2(1, "2.2", "0.05", 1, "2.2", "0.05", 1)
End If


OS_Result = 0
rv0 = 0

CardResult = DO_WritePort(card, Channel_P1C, &H0)
                 
MsecDelay (0.3)

OpenShortTest_Result

If OS_Result <> 1 Then
    rv0 = 0                 'OS Fail
    GoTo AU6371ELResult
End If

CardResult = DO_WritePort(card, Channel_P1C, &HFF)
                       
      
  Tester.Print "Begin AU6433EF FT2 Test: NB mode test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim CPRMMODE As Byte
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
               
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
                 Call MsecDelay(0.1)
                 
                ' CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                ' Call MsecDelay(1#)    'power on time          'NBMD
                ChipString = "vid"
                ' If GetDeviceName(ChipString) <> "" Then
                '    rv0 = 3
                '    GoTo AU6371ELResult
                '
                '  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
               '   Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1.3)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      
                      
                      If rv0 = 1 Then
                        
                        If Dir("D:\LABPC.PC") = "LABPC.PC" Then
                            Call PowerSet2(1, "3.3", "0.05", 1, "1.8", "0.05", 1)
                        Else
                            Call PowerSet2(1, "1.8", "0.05", 1, "1.8", "0.05", 1)
                        End If
                        
                        rv0 = CBWTest_New(0, 1, ChipString)
                      End If
                      
                      If rv0 = 1 Then
                           rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                           If rv0 <> 1 Then
                           rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                       End If
                      
                       If rv0 <> 1 Then ' for E55 command
                          rv0 = Read_SD_SpeedE55(0, 0, 64, "8Bits")
                          If rv0 <> 1 Then
                          rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                          If rv0 = 1 Then
                            CPRMMODE = 1
                          End If
                      End If
                      ClosePipe
                      
                      If (rv0 = 0) Or (rv0 = 2) Then 'for OSE request
                        rv0 = 3
                        GoTo AU6371ELResult
                      End If
                      
                      Tester.Print "rv0="; rv0
                     
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                          Tester.Print "LightON="; LightOn
                          Tester.Print "LightOFF="; LightOff
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                    
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      If rv0 = 1 Then
                        TmpLBA = LBA
                        'LBA = 99
                         'For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6371ELResult
                             End If
                         'Next
                        LBA = TmpLBA
                      End If
'=======================================================================================
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
                 Call MsecDelay(0.1)
               
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
                 Call MsecDelay(0.1)
                  OpenPipe
                  rv5 = ReInitial(0)
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                
                If CPRMMODE = 0 Then  ' for E54 Before
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                 Else             ' for AU6433E55 after
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476E55(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                End If
                
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H7F)  'Disconnect all card check NBMD
                Call MsecDelay(0.2)
                
                If GetDeviceName(ChipString) <> "" Then
                    rv0 = 3
                End If
                
                
                
AU6371ELResult:
                        CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power

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
Public Sub AU6433BLF2DTestSub()


'201012/30 This code just for V2 S/B
'reduce NBMD test time
'2011/3/24 modify LV: 1.8 -> 1.75 , add MSpro R/W 128 Sector
'2011/3/30 modify MPpro & XD R/W 2K => 4K

Dim TmpLBA As Long
Dim i As Integer

               If PCI7248InitFinish = 0 Then
                  PCI7248ExistAU6254
                  Call SetTimer_1ms
               End If
If Dir("D:\LABPC.PC") = "LABPC.PC" Then
    Call PowerSet2(1, "3.3", "0.05", 1, "2.2", "0.05", 1)
Else
    Call PowerSet2(1, "2.2", "0.05", 1, "2.2", "0.05", 1)
End If


OS_Result = 0
rv0 = 0

CardResult = DO_WritePort(card, Channel_P1C, &H0)
                 
MsecDelay (0.3)

OpenShortTest_Result

If OS_Result <> 1 Then
    rv0 = 0                 'OS Fail
    GoTo AU6371ELResult
End If

CardResult = DO_WritePort(card, Channel_P1C, &HFF)
                       
      
  Tester.Print "Begin AU6433DL FT2 Test:"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim CPRMMODE As Byte
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
               
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
                 Call MsecDelay(0.1)
                 
                ' CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                ' Call MsecDelay(1#)    'power on time          'NBMD
                ChipString = "vid"
                ' If GetDeviceName(ChipString) <> "" Then
                '    rv0 = 3
                '    GoTo AU6371ELResult
                '
                '  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
               '   Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1.3)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      
                      
                      If rv0 = 1 Then
                        
                        If Dir("D:\LABPC.PC") = "LABPC.PC" Then
                            Call PowerSet2(1, "3.3", "0.05", 1, "1.75", "0.05", 1)
                        Else
                            Call PowerSet2(1, "1.75", "0.05", 1, "1.75", "0.05", 1)
                        End If
                        
                        rv0 = CBWTest_New(0, 1, ChipString)
                      End If
                      
                      If rv0 = 1 Then
                           rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                           If rv0 <> 1 Then
                           rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                       End If
                      
                       If rv0 <> 1 Then ' for E55 command
                          rv0 = Read_SD_SpeedE55(0, 0, 64, "8Bits")
                          If rv0 <> 1 Then
                          rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                          If rv0 = 1 Then
                            CPRMMODE = 1
                          End If
                      End If
                      ClosePipe
                      
                      If (rv0 = 0) Or (rv0 = 2) Then 'for OSE request
                        rv0 = 3
                        GoTo AU6371ELResult
                      End If
                      
                      Tester.Print "rv0="; rv0
                     
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                          Tester.Print "LightON="; LightOn
                          Tester.Print "LightOFF="; LightOff
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                    
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      If rv0 = 1 Then
                        TmpLBA = LBA
                        'LBA = 99
                         'For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6371ELResult
                             End If
                         'Next
                        LBA = TmpLBA
                      End If
'=======================================================================================
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
                 Call MsecDelay(0.1)
               
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
                 Call MsecDelay(0.1)
                  OpenPipe
                  rv5 = ReInitial(0)
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                
                If CPRMMODE = 0 Then  ' for E54 Before
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                 Else             ' for AU6433E55 after
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476E55(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                End If
                
                'If rv5 = 1 Then
                '    rv5 = CBWTest_New_128_Sector_AU6377(0, 1)
                'End If
                
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H7F)  'Disconnect all card check NBMD
                Call MsecDelay(0.2)
                
                If GetDeviceName(ChipString) <> "" Then
                    rv0 = 3
                End If
                
                
                
AU6371ELResult:
                        CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power

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
Public Sub AU6433BLF30TestSub()
Dim TmpLBA As Long
Dim i As Integer
Dim Serch_Dev As Byte
Dim Serch_Dev_Count As Integer


               If PCI7248InitFinish = 0 Then
                  PCI7248ExistAU6254
                  Call SetTimer_1ms
               End If

Call PowerSet2(1, "3.3", "0.05", 1, "2.1", "0.05", 1)

OS_Result = 0
rv0 = 0

CardResult = DO_WritePort(card, Channel_P1C, &H0)
                 
MsecDelay (0.1)

OpenShortTest_Result

If OS_Result <> 1 Then
    rv0 = 0                 'OS Fail
    GoTo AU6371ELResult
End If

CardResult = DO_WritePort(card, Channel_P1C, &HFF)
      
  Tester.Print "AU6433EF : NB mode test , change 1.8V to 2.1V"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim CPRMMODE As Byte
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(0.2)    'power on time
                 ChipString = "vid"
                 
                 If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
               '   Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     
                     Call MsecDelay(1.3)
                     
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
                      
                       If rv0 <> 1 Then ' for E55 command
                          rv0 = Read_SD_SpeedE55(0, 0, 64, "8Bits")
                          If rv0 <> 1 Then
                          rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                          If rv0 = 1 Then
                            CPRMMODE = 1
                          End If
                      End If
                      ClosePipe
                      
                      Tester.Print "rv0="; rv0
                     
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                          Tester.Print "LightON="; LightOn
                          Tester.Print "LightOFF="; LightOff
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                    
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      
                     TmpLBA = LBA
                     'LBA = 99
                         'For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6371ELResult
                             End If
                         'Next
                      LBA = TmpLBA
                      
'=======================================================================================
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  'CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F) 'SD off
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.2)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
                  Call MsecDelay(0.4)
                  
                Do Until (Serch_Dev = 1) Or (Serch_Dev_Count > 10)
                    If GetDeviceName(ChipString) <> "" Then
                        Serch_Dev = 1
                    End If
                    
                    Serch_Dev_Count = Serch_Dev_Count + 1
                    Call MsecDelay(0.1)
                Loop
                 
                  'OpenPipe
                  'rv3 = ReInitial(0)
                
                 
                'ClosePipe
                
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  MS Card test
                '================================================
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
                'CardResult = DO_WritePort(card, Channel_P1A, &H7F)
              
              
                 Call MsecDelay(0.01)
               
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
                 'Call MsecDelay(1.2)
                  OpenPipe
                  rv5 = ReInitial(0)
                
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                
                If CPRMMODE = 0 Then  ' for E54 Before
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                 Else             ' for AU6433E55 after
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476E55(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
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
Public Sub AU6433BLF3BTestSub()


'201012/30 This code just for V5 S/B

Dim TmpLBA As Long
Dim i As Integer

               If PCI7248InitFinish = 0 Then
                  PCI7248ExistAU6254
                  Call SetTimer_1ms
               End If

Call PowerSet2(1, "2.2", "0.05", 1, "2.2", "0.05", 1)

OS_Result = 0
rv0 = 0

CardResult = DO_WritePort(card, Channel_P1C, &H0)
                 
MsecDelay (0.3)

OpenShortTest_Result

If OS_Result <> 1 Then
    rv0 = 0                 'OS Fail
    GoTo AU6371ELResult
End If

CardResult = DO_WritePort(card, Channel_P1C, &HFF)
                       
      
  Tester.Print "Begin AU6433EF FT2 Test: NB mode test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim CPRMMODE As Byte
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
               
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
                 Call MsecDelay(0.1)
                 
                ' CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                ' Call MsecDelay(1#)    'power on time          'NBMD
                ChipString = "vid"
                ' If GetDeviceName(ChipString) <> "" Then
                '    rv0 = 3
                '    GoTo AU6371ELResult
                '
                '  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
               '   Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1.3)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      
                      
                      If rv0 = 1 Then
                        'CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                        'Call MsecDelay(0.2)
                        Call PowerSet2(1, "1.8", "0.05", 1, "1.8", "0.05", 1)
                        'CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                        'Call MsecDelay(1.3)
                        
                        rv0 = CBWTest_New(0, 1, ChipString)
                      End If
                      
                      If rv0 = 1 Then
                           rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                           If rv0 <> 1 Then
                           rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                       End If
                      
                       If rv0 <> 1 Then ' for E55 command
                          rv0 = Read_SD_SpeedE55(0, 0, 64, "8Bits")
                          If rv0 <> 1 Then
                          rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                          If rv0 = 1 Then
                            CPRMMODE = 1
                          End If
                      End If
                      ClosePipe
                      
                      If rv0 = 0 Then   'for OSE request
                        rv0 = 3
                      End If
                      
                      Tester.Print "rv0="; rv0
                     
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                          Tester.Print "LightON="; LightOn
                          Tester.Print "LightOFF="; LightOff
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                    
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      If rv0 = 1 Then
                        TmpLBA = LBA
                        'LBA = 99
                         'For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6371ELResult
                             End If
                         'Next
                        LBA = TmpLBA
                      End If
'=======================================================================================
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
                 Call MsecDelay(0.1)
               
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
                 Call MsecDelay(0.1)
                  OpenPipe
                  rv5 = ReInitial(0)
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                
                If CPRMMODE = 0 Then  ' for E54 Before
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                 Else             ' for AU6433E55 after
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476E55(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                End If
                
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H7F)  'Disconnect all card check NBMD
                Call MsecDelay(0.2)
                
                If GetDeviceName(ChipString) <> "" Then
                    rv0 = 3
                End If
                
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
Public Sub AU6433BLF3CTestSub()


'201012/30 This code just for V5 S/B
'reduce NBMD test time

Dim TmpLBA As Long
Dim i As Integer

               If PCI7248InitFinish = 0 Then
                  PCI7248ExistAU6254
                  Call SetTimer_1ms
               End If
If Dir("D:\LABPC.PC") = "LABPC.PC" Then
    Call PowerSet2(1, "3.3", "0.05", 1, "2.2", "0.05", 1)
Else
    Call PowerSet2(1, "2.2", "0.05", 1, "2.2", "0.05", 1)
End If

OS_Result = 0
rv0 = 0

CardResult = DO_WritePort(card, Channel_P1C, &H0)
                 
MsecDelay (0.3)

OpenShortTest_Result

If OS_Result <> 1 Then
    rv0 = 0                 'OS Fail
    GoTo AU6371ELResult
End If

CardResult = DO_WritePort(card, Channel_P1C, &HFF)
                       
      
  Tester.Print "Begin AU6433EF FT2 Test: NB mode test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim CPRMMODE As Byte
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
               
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
                 Call MsecDelay(0.1)
                 
                ' CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                ' Call MsecDelay(1#)    'power on time          'NBMD
                ChipString = "vid"
                ' If GetDeviceName(ChipString) <> "" Then
                '    rv0 = 3
                '    GoTo AU6371ELResult
                '
                '  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
               '   Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1.3)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      
                      
                      If rv0 = 1 Then
                        
                        If Dir("D:\LABPC.PC") = "LABPC.PC" Then
                            Call PowerSet2(1, "3.3", "0.05", 1, "1.8", "0.05", 1)
                        Else
                            Call PowerSet2(1, "1.8", "0.05", 1, "1.8", "0.05", 1)
                        End If
                        
                        rv0 = CBWTest_New(0, 1, ChipString)
                      End If
                      
                      If rv0 = 1 Then
                           rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                           If rv0 <> 1 Then
                           rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                       End If
                      
                       If rv0 <> 1 Then ' for E55 command
                          rv0 = Read_SD_SpeedE55(0, 0, 64, "8Bits")
                          If rv0 <> 1 Then
                          rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                          If rv0 = 1 Then
                            CPRMMODE = 1
                          End If
                      End If
                      ClosePipe
                      
                      If (rv0 = 0) Or (rv0 = 2) Then 'for OSE request
                        rv0 = 2
                        GoTo AU6371ELResult
                      End If
                      
                      Tester.Print "rv0="; rv0
                     
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                          Tester.Print "LightON="; LightOn
                          Tester.Print "LightOFF="; LightOff
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                    
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      If rv0 = 1 Then
                        TmpLBA = LBA
                        'LBA = 99
                         'For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6371ELResult
                             End If
                         'Next
                        LBA = TmpLBA
                      End If
'=======================================================================================
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
                 Call MsecDelay(0.1)
               
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
                 Call MsecDelay(0.1)
                  OpenPipe
                  rv5 = ReInitial(0)
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                
                If CPRMMODE = 0 Then  ' for E54 Before
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                 Else             ' for AU6433E55 after
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476E55(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                End If
                
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H7F)  'Disconnect all card check NBMD
                Call MsecDelay(0.2)
                
                If GetDeviceName(ChipString) <> "" Then
                    rv0 = 3
                End If
                
                
                
AU6371ELResult:

                CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
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
Public Sub AU6433BLF3DTestSub()


'201012/30 This code just for V5 S/B
'reduce NBMD test time
'2011/3/24 modify LV: 1.8 -> 1.75 , add MSpro R/W 128 Sector
'2011/3/30 modify MPpro & XD R/W 2K => 4K

Dim TmpLBA As Long
Dim i As Integer

               If PCI7248InitFinish = 0 Then
                  PCI7248ExistAU6254
                  Call SetTimer_1ms
               End If
If Dir("D:\LABPC.PC") = "LABPC.PC" Then
    Call PowerSet2(1, "3.3", "0.05", 1, "2.2", "0.05", 1)
Else
    Call PowerSet2(1, "2.2", "0.05", 1, "2.2", "0.05", 1)
End If


OS_Result = 0
rv0 = 0

CardResult = DO_WritePort(card, Channel_P1C, &H0)
                 
MsecDelay (0.3)

OpenShortTest_Result

If OS_Result <> 1 Then
    rv0 = 0                 'OS Fail
    GoTo AU6371ELResult
End If

CardResult = DO_WritePort(card, Channel_P1C, &HFF)
                       
      
  Tester.Print "Begin AU6433DL FT2 Test: "
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim CPRMMODE As Byte
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
               
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
                 Call MsecDelay(0.1)
                 
                ' CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                ' Call MsecDelay(1#)    'power on time          'NBMD
                ChipString = "vid"
                ' If GetDeviceName(ChipString) <> "" Then
                '    rv0 = 3
                '    GoTo AU6371ELResult
                '
                '  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
               '   Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1.3)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                      ClosePipe
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      
                      If rv0 = 1 Then
                        
                        If Dir("D:\LABPC.PC") = "LABPC.PC" Then
                            Call PowerSet2(1, "3.3", "0.05", 1, "1.75", "0.05", 1)
                        Else
                            Call PowerSet2(1, "1.75", "0.05", 1, "1.75", "0.05", 1)
                        End If
                        
                        rv0 = CBWTest_New(0, 1, ChipString)
                      End If
                      
                      If rv0 = 1 Then
                           rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                           If rv0 <> 1 Then
                           rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                       End If
                      
                       If rv0 <> 1 Then ' for E55 command
                          rv0 = Read_SD_SpeedE55(0, 0, 64, "8Bits")
                          If rv0 <> 1 Then
                          rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                          If rv0 = 1 Then
                            CPRMMODE = 1
                          End If
                      End If
                      ClosePipe
                      
                      If (rv0 = 0) Or (rv0 = 2) Then 'for OSE request
                        rv0 = 3
                        GoTo AU6371ELResult
                      End If
                      
                      Tester.Print "rv0="; rv0
                     
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                          Tester.Print "LightON="; LightOn
                          Tester.Print "LightOFF="; LightOff
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                    
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      If rv0 = 1 Then
                        TmpLBA = LBA
                        'LBA = 99
                         'For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6371ELResult
                             End If
                         'Next
                        LBA = TmpLBA
                      End If
'=======================================================================================
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
                 Call MsecDelay(0.1)
               
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
                 Call MsecDelay(0.1)
                  OpenPipe
                  rv5 = ReInitial(0)
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                
                If CPRMMODE = 0 Then  ' for E54 Before
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                 Else             ' for AU6433E55 after
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476E55(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                End If
                
                'If rv5 = 1 Then
                '    rv5 = CBWTest_New_128_Sector_AU6377(0, 1)
                'End If
                
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H7F)  'Disconnect all card check NBMD
                Call MsecDelay(0.2)
                
                If GetDeviceName(ChipString) <> "" Then
                    rv0 = 3
                End If
                
                
                
AU6371ELResult:
                        CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power

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

Public Sub AU6433BLF3FTestSub()


'201012/30 This code just for V5 S/B
'reduce NBMD test time
'2011/3/24 modify LV: 1.8 -> 1.75 , add MSpro R/W 128 Sector
'2011/3/30 modify MSpro & XD R/W 2K => 4K
'2011/11/30 add MSpro R/W 64K

Dim TmpLBA As Long
Dim i As Integer

               If PCI7248InitFinish = 0 Then
                  PCI7248ExistAU6254
                  Call SetTimer_1ms
               End If
If Dir("D:\LABPC.PC") = "LABPC.PC" Then
    Call PowerSet2(1, "3.3", "0.05", 1, "2.2", "0.05", 1)
Else
    Call PowerSet2(1, "2.2", "0.05", 1, "2.2", "0.05", 1)
End If


OS_Result = 0
rv0 = 0

CardResult = DO_WritePort(card, Channel_P1C, &H0)
                 
MsecDelay (0.3)

OpenShortTest_Result

If OS_Result <> 1 Then
    rv0 = 0                 'OS Fail
    GoTo AU6371ELResult
End If

CardResult = DO_WritePort(card, Channel_P1C, &HFF)
                       
      
  Tester.Print "Begin AU6433DL FT2 Test: "
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim CPRMMODE As Byte
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
               
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
                 Call MsecDelay(0.1)
                 
                ' CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                ' Call MsecDelay(1#)    'power on time          'NBMD
                ChipString = "vid"
                ' If GetDeviceName(ChipString) <> "" Then
                '    rv0 = 3
                '    GoTo AU6371ELResult
                '
                '  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
               '   Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1.3)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                      ClosePipe
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
                      
                      If rv0 = 1 Then
                        
                        If Dir("D:\LABPC.PC") = "LABPC.PC" Then
                            Call PowerSet2(1, "3.3", "0.05", 1, "1.75", "0.05", 1)
                        Else
                            Call PowerSet2(1, "1.75", "0.05", 1, "1.75", "0.05", 1)
                        End If
                        
                        rv0 = CBWTest_New(0, 1, ChipString)
                      End If
                      
                      If rv0 = 1 Then
                           rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                           If rv0 <> 1 Then
                           rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                       End If
                      
                       If rv0 <> 1 Then ' for E55 command
                          rv0 = Read_SD_SpeedE55(0, 0, 64, "8Bits")
                          If rv0 <> 1 Then
                          rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                          If rv0 = 1 Then
                            CPRMMODE = 1
                          End If
                      End If
                      ClosePipe
                      
                      If (rv0 = 0) Then 'for OSE request
                        rv0 = 2
                        'GoTo AU6371ELResult
                      End If
                      
                      Tester.Print "rv0="; rv0
                     
                        If rv0 = 1 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                          Tester.Print "LightON="; LightOn
                          Tester.Print "LightOFF="; LightOff
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                    
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      If rv0 = 1 Then
                        TmpLBA = LBA
                        'LBA = 99
                         'For i = 1 To 30
                             LBA = LBA + 199
                            
                             'ClosePipe
                             rv0 = CBWTest_New_128_Sector_AU6377(0, rv0)  ' write
                             ClosePipe
                             'If rv1 <> 1 Then
                             '   LBA = TmpLBA
                             '   GoTo AU6371ELResult
                             'End If
                         'Next
                        LBA = TmpLBA
                      End If
'=======================================================================================
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
                 Call MsecDelay(0.1)
               
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
                 Call MsecDelay(0.1)
                  OpenPipe
                  rv5 = ReInitial(0)
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                
                If CPRMMODE = 0 Then  ' for E54 Before
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                 Else             ' for AU6433E55 after
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476E55(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                End If
                ClosePipe
                
                If rv5 = 1 Then
                    rv5 = CBWTest_New_128_Sector_AU6377(0, rv5)
                    ClosePipe
                End If
                
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                'ClosePipe
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H7F)  'Disconnect all card check NBMD
                Call MsecDelay(0.2)
                
                If GetDeviceName(ChipString) <> "" Then
                    rv0 = 3
                End If
                
                
                
AU6371ELResult:
                        CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power

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

Public Sub AU6433BLF3ETestSub()

'2011/6/13 Just modify O/S pattern (pin2:open,pin11:short)
'201012/30 This code just for V5 S/B
'reduce NBMD test time
'2011/3/24 modify LV: 1.8 -> 1.75 , add MSpro R/W 128 Sector
'2011/3/30 modify MPpro & XD R/W 2K => 4K

Dim TmpLBA As Long
Dim i As Integer

    If PCI7248InitFinish = 0 Then
        PCI7248ExistAU6254
        Call SetTimer_1ms
    End If
               
If Dir("D:\LABPC.PC") = "LABPC.PC" Then
    Call PowerSet2(1, "3.3", "0.05", 1, "2.2", "0.05", 1)
Else
    Call PowerSet2(1, "2.2", "0.05", 1, "2.2", "0.05", 1)
End If


OS_Result = 0
rv0 = 0

CardResult = DO_WritePort(card, Channel_P1C, &H0)
                 
MsecDelay (0.3)

OpenShortTest_Result

If OS_Result <> 1 Then
    rv0 = 0                 'OS Fail
    GoTo AU6371ELResult
End If

CardResult = DO_WritePort(card, Channel_P1C, &HFF)
                       
      
  Tester.Print "Begin AU6433DL FT2 Test: "
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim CPRMMODE As Byte
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
               
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
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                    If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 Call MsecDelay(0.2)
                 
                ' CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                ' Call MsecDelay(1#)    'power on time          'NBMD
                ChipString = "vid"
                ' If GetDeviceName(ChipString) <> "" Then
                '    rv0 = 3
                '    GoTo AU6371ELResult
                '
                '  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
               '   Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     'Call MsecDelay(1.3)
                     
                     rv0 = WaitDevOn(ChipString)
                     Call MsecDelay(0.1)
                     
                     
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                      ClosePipe
                      
                      If rv0 = 1 Then
                        rv0 = CBWTest_New(0, 1, ChipString)
                      End If
                      
                      If rv0 = 1 Then
                        
                        If Dir("D:\LABPC.PC") = "LABPC.PC" Then
                            Call PowerSet2(1, "3.3", "0.05", 1, "1.75", "0.05", 1)
                        Else
                            Call PowerSet2(1, "1.75", "0.05", 1, "1.75", "0.05", 1)
                        End If
                        
                        rv0 = CBWTest_New(0, 1, ChipString)
                      End If
                      
                      If rv0 = 1 Then
                           rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                           If rv0 <> 1 Then
                           rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                       End If
                      
                       If rv0 <> 1 Then ' for E55 command
                          rv0 = Read_SD_SpeedE55(0, 0, 64, "8Bits")
                          If rv0 <> 1 Then
                          rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                          If rv0 = 1 Then
                            CPRMMODE = 1
                          End If
                      End If
                      ClosePipe
                      
                      If (rv0 = 0) Or (rv0 = 2) Then 'for OSE request
                        rv0 = 3
                        GoTo AU6371ELResult
                      End If
                      
                      'Tester.Print "rv0="; rv0
                     
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                          Tester.Print "LightON="; LightOn
                          Tester.Print "LightOFF="; LightOff
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                    
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      If rv0 = 1 Then
                        TmpLBA = LBA
                        'LBA = 99
                         'For i = 1 To 30
                             'rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv0 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             'If rv0 <> 1 Then
                             '   LBA = TmpLBA
                             '   GoTo AU6371ELResult
                             'End If
                         'Next
                        LBA = TmpLBA
                      End If
'=======================================================================================
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
                 Call MsecDelay(0.1)
               
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
                 Call MsecDelay(0.1)
                  OpenPipe
                  rv5 = ReInitial(0)
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                
                If CPRMMODE = 0 Then  ' for E54 Before
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                 Else             ' for AU6433E55 after
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476E55(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                End If
                
                'If rv5 = 1 Then
                '    rv5 = CBWTest_New_128_Sector_AU6377(0, 1)
                'End If
                
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H7F)  'Disconnect all card check NBMD
                Call MsecDelay(0.2)
                
                If GetDeviceName(ChipString) <> "" Then
                    rv0 = 3
                End If
                
                
                
AU6371ELResult:
                        CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power

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
Public Sub AU6433BLF2ETestSub()

'2011/6/16 AU6433S61-RBL FT2 test
'2011/6/13 Just modify O/S pattern (pin2:open,pin11:short)
'201012/30 This code just for V5 S/B
'reduce NBMD test time
'2011/3/24 modify LV: 1.8 -> 1.75 , add MSpro R/W 128 Sector
'2011/3/30 modify MPpro & XD R/W 2K => 4K

Dim TmpLBA As Long
Dim i As Integer

If PCI7248InitFinish = 0 Then
    PCI7248ExistAU6254
    Call SetTimer_1ms
End If

If Dir("D:\LABPC.PC") = "LABPC.PC" Then
    Call PowerSet2(1, "3.3", "0.05", 1, "2.2", "0.05", 1)
Else
    Call PowerSet2(1, "2.2", "0.05", 1, "2.2", "0.05", 1)
End If


'OS_Result = 0
'rv0 = 0

'CardResult = DO_WritePort(card, Channel_P1C, &H0)
                 
'MsecDelay (0.3)

'OpenShortTest_Result

'If OS_Result <> 1 Then
'    rv0 = 0                 'OS Fail
'    GoTo AU6371ELResult
'End If

'CardResult = DO_WritePort(card, Channel_P1C, &HFF)
                       
      
  Tester.Print "Begin AU6433DL FT2 Test: "
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim CPRMMODE As Byte
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
               
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
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                    If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 Call MsecDelay(0.2)
                 
                ' CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                ' Call MsecDelay(1#)    'power on time          'NBMD
                ChipString = "vid"
                ' If GetDeviceName(ChipString) <> "" Then
                '    rv0 = 3
                '    GoTo AU6371ELResult
                '
                '  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
               '   Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     'Call MsecDelay(1.3)
                     
                     rv0 = WaitDevOn(ChipString)
                     Call MsecDelay(0.1)
                     
                     
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                      ClosePipe
                      
                      If rv0 = 1 Then
                        rv0 = CBWTest_New(0, 1, ChipString)
                      End If
                      
                      If rv0 = 1 Then
                        
                        If Dir("D:\LABPC.PC") = "LABPC.PC" Then
                            Call PowerSet2(1, "3.3", "0.05", 1, "1.75", "0.05", 1)
                        Else
                            Call PowerSet2(1, "1.75", "0.05", 1, "1.75", "0.05", 1)
                        End If
                        
                        rv0 = CBWTest_New(0, 1, ChipString)
                      End If
                      
                      If rv0 = 1 Then
                           rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                           If rv0 <> 1 Then
                           rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                       End If
                      
                       If rv0 <> 1 Then ' for E55 command
                          rv0 = Read_SD_SpeedE55(0, 0, 64, "8Bits")
                          If rv0 <> 1 Then
                          rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                          If rv0 = 1 Then
                            CPRMMODE = 1
                          End If
                      End If
                      ClosePipe
                      
                      If (rv0 = 0) Or (rv0 = 2) Then 'for OSE request
                        rv0 = 3
                        GoTo AU6371ELResult
                      End If
                      
                      'Tester.Print "rv0="; rv0
                     
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                          Tester.Print "LightON="; LightOn
                          Tester.Print "LightOFF="; LightOff
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                    
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      If rv0 = 1 Then
                        TmpLBA = LBA
                        'LBA = 99
                         'For i = 1 To 30
                             'rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv0 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             'If rv0 <> 1 Then
                             '   LBA = TmpLBA
                             '   GoTo AU6371ELResult
                             'End If
                         'Next
                        LBA = TmpLBA
                      End If
'=======================================================================================
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
                 Call MsecDelay(0.1)
               
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
                 Call MsecDelay(0.1)
                  OpenPipe
                  rv5 = ReInitial(0)
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                
                If CPRMMODE = 0 Then  ' for E54 Before
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                 Else             ' for AU6433E55 after
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476E55(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                End If
                
                'If rv5 = 1 Then
                '    rv5 = CBWTest_New_128_Sector_AU6377(0, 1)
                'End If
                
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H7F)  'Disconnect all card check NBMD
                Call MsecDelay(0.2)
                
                If GetDeviceName(ChipString) <> "" Then
                    rv0 = 3
                End If
                
                
                
AU6371ELResult:
                        CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power

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

Public Sub AU6433BLF2FTestSub()

'2011/6/16 AU6433S61-RBL FT2 test
'2011/6/13 Just modify O/S pattern (pin2:open,pin11:short)
'201012/30 This code just for V5 S/B
'reduce NBMD test time
'2011/3/24 modify LV: 1.8 -> 1.75 , add MSpro R/W 128 Sector
'2011/3/30 modify MSpro & XD R/W 2K => 4K
'2011/11/30 add MSpro R/W 64K

Dim TmpLBA As Long
Dim i As Integer

If PCI7248InitFinish = 0 Then
    PCI7248ExistAU6254
    Call SetTimer_1ms
End If

If Dir("D:\LABPC.PC") = "LABPC.PC" Then
    Call PowerSet2(1, "3.3", "0.05", 1, "2.2", "0.05", 1)
Else
    Call PowerSet2(1, "2.2", "0.05", 1, "2.2", "0.05", 1)
End If


'OS_Result = 0
'rv0 = 0

'CardResult = DO_WritePort(card, Channel_P1C, &H0)
                 
'MsecDelay (0.3)

'OpenShortTest_Result

'If OS_Result <> 1 Then
'    rv0 = 0                 'OS Fail
'    GoTo AU6371ELResult
'End If

'CardResult = DO_WritePort(card, Channel_P1C, &HFF)
                       
      
  Tester.Print "Begin AU6433DL FT2 Test: "
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim CPRMMODE As Byte
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
               
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
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                    If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 Call MsecDelay(0.2)
                 
                ' CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                ' Call MsecDelay(1#)    'power on time          'NBMD
                ChipString = "vid"
                ' If GetDeviceName(ChipString) <> "" Then
                '    rv0 = 3
                '    GoTo AU6371ELResult
                '
                '  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
               '   Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     'Call MsecDelay(1.3)
                     
                     rv0 = WaitDevOn(ChipString)
                     Call MsecDelay(0.1)
                     
                     
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                      ClosePipe
                      
                      If rv0 = 1 Then
                        rv0 = CBWTest_New(0, 1, ChipString)
                        ClosePipe
                      End If
                      
                      If rv0 = 1 Then
                        
                        If Dir("D:\LABPC.PC") = "LABPC.PC" Then
                            Call PowerSet2(1, "3.3", "0.05", 1, "1.75", "0.05", 1)
                        Else
                            Call PowerSet2(1, "1.75", "0.05", 1, "1.75", "0.05", 1)
                        End If
                        
                        rv0 = CBWTest_New(0, 1, ChipString)
                      End If
                      
                      If rv0 = 1 Then
                           rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                           If rv0 <> 1 Then
                           rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                       End If
                      
                       If rv0 <> 1 Then ' for E55 command
                          rv0 = Read_SD_SpeedE55(0, 0, 64, "8Bits")
                          If rv0 <> 1 Then
                          rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                          If rv0 = 1 Then
                            CPRMMODE = 1
                          End If
                      End If
                      ClosePipe
                      
                      
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                          Tester.Print "LightON="; LightOn
                          Tester.Print "LightOFF="; LightOff
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                    
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      If rv0 = 1 Then
                        TmpLBA = LBA
                        'LBA = 99
                         'For i = 1 To 30
                             'rv1 = 0
                             LBA = LBA + 199
                            
                             'ClosePipe
                             rv0 = CBWTest_New_128_Sector_AU6377(0, rv0)  ' write
                             ClosePipe
                             'If rv0 <> 1 Then
                             '   LBA = TmpLBA
                             '   GoTo AU6371ELResult
                             'End If
                         'Next
                        LBA = TmpLBA
                      End If
'=======================================================================================
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
                 Call MsecDelay(0.1)
               
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
                 Call MsecDelay(0.1)
                  OpenPipe
                  rv5 = ReInitial(0)
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                
                If CPRMMODE = 0 Then  ' for E54 Before
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                 Else             ' for AU6433E55 after
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476E55(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                End If
                
                ClosePipe
                
                If rv5 = 1 Then
                    'ClosePipe
                    rv5 = CBWTest_New_128_Sector_AU6377(0, rv5)
                    ClosePipe
                End If
                
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                'ClosePipe
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H7F)  'Disconnect all card check NBMD
                Call MsecDelay(0.2)
                
                If GetDeviceName(ChipString) <> "" Then
                    rv0 = 3
                End If
                
                
                
AU6371ELResult:
                        CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power

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

Public Sub AU6433BLFEETestSub()

'2011/6/16 AU6433S61-RBL FT2 test
'2011/6/13 Just modify O/S pattern (pin2:open,pin11:short)
'201012/30 This code just for V5 S/B
'reduce NBMD test time
'2011/3/24 modify LV: 1.8 -> 1.75 , add MSpro R/W 128 Sector
'2011/3/30 modify MPpro & XD R/W 2K => 4K
'2011/10/3 for all voltage scan eng V33: 3.6 ~ 3.0 ; V18: 2.2 ~ 1.6


Dim TmpLBA As Long
Dim i As Integer
Dim V33_Count As Integer
Dim V18_Count As Integer

Dim CurV33 As String
Dim CurV18 As String


    CurV33 = "3.6"
    CurV18 = "2.2"
    V33_Count = 0
    V18_Count = 0

    If PCI7248InitFinish = 0 Then
        PCI7248ExistAU6254
        Call SetTimer_1ms
    End If
      
    Tester.Print "Begin AU6433DL FT2 Test: "

    Dim ChipString As String
    Dim AU6371EL_SD As Byte
    Dim AU6371EL_CF As Byte
    Dim AU6371EL_XD As Byte
    Dim AU6371EL_MS As Byte
    Dim AU6371EL_MSP  As Byte
    Dim AU6371EL_BootTime As Single
    Dim CPRMMODE As Byte
    OldChipName = ""
                 
    ' initial condition
                
    AU6371EL_SD = 1
    AU6371EL_CF = 2
    AU6371EL_XD = 8
    AU6371EL_MS = 32
    AU6371EL_MSP = 64
    AU6371EL_BootTime = 0.6
               
               
Routine_Loop:
    
    CurV33 = CStr(3.6 - (CSng(V33_Count) / 10))
    CurV18 = CStr(2.2 - (CSng(V18_Count) / 10))
    Call PowerSet2(1, CurV33, "0.3", 1, CurV18, "0.3", 1)
    Tester.Cls
    
    Tester.Print "V33= "; CurV33 & vbTab & "V18= "; CurV18
            
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
    ChipString = "vid"
    
    
    '=========================================
    '    POWER on
    '=========================================
    
    CardResult = DO_WritePort(card, Channel_P1A, &H7F)
    If CardResult <> 0 Then
        MsgBox "Power off fail"
        End
    End If
        
    Call MsecDelay(0.2)
        
    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
    If CardResult <> 0 Then
        MsgBox "Read light off fail"
        End
    End If
               
                
    '===============================================
    '  SD Card test
    '===============================================
              
    ' set SD card detect down
    CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
    If CardResult <> 0 Then
        MsgBox "Set SD Card Detect Down Fail"
        End
    End If
                     
    rv0 = WaitDevOn(ChipString)
    Call MsecDelay(0.1)
                     
                     
    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
    If CardResult <> 0 Then
        MsgBox "Read light On fail"
        End
    End If
                     
    ClosePipe
                      
    If rv0 = 1 Then
        rv0 = CBWTest_New(0, 1, ChipString)
    End If
                      
    If rv0 = 1 Then
        rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
        If rv0 <> 1 Then
            rv0 = 2
            Tester.Print "SD bus width Fail"
        End If
    End If
                      
    ClosePipe
                     
    If rv0 <> 0 Then
        If LightOn <> &HBF Or LightOff <> &HFF Then
            Tester.Print "LightON="; LightOn
            Tester.Print "LightOFF="; LightOff
            UsbSpeedTestResult = GPO_FAIL
            rv0 = 3
        End If
    End If
                                      
    
    '=======================================================================================
    '       SD R/W 64K
    '=======================================================================================
    If rv0 = 1 Then
        rv0 = CBWTest_New_128_Sector_AU6377(0, rv0)  ' write
        If rv0 <> 1 Then
            Tester.Print "128 Sector R/W Fail ..."
        End If
    End If
    ClosePipe
    Call LabelMenu(0, rv0, 1)   ' no card test fail
    Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    '===============================================
    '  CF Card test
    '================================================
               
    rv1 = rv0  '----------- no CF slot
    'Call LabelMenu(1, rv1, rv0)
    'Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    
    '===============================================
    '  SMC Card test  : stop these test for card not enough
    '================================================
              
     rv2 = rv1   'AU6433 has no SMC slot
           
    
    '===============================================
    '  XD Card test
    '================================================
    CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                     
    If CardResult <> 0 Then
        MsgBox "Set XD Card Detect On Fail"
        End
    End If
                 
    Call MsecDelay(0.1)
    CardResult = DO_WritePort(card, Channel_P1A, &H77)
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
                   
    rv4 = rv3  'AU6344 has no MS slot pin
               
    'Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    '===============================================
    '  MS Pro Card test
    '================================================
              
    CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
    Call MsecDelay(0.1)
    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
    Call MsecDelay(0.1)
    OpenPipe
    rv5 = ReInitial(0)
    ClosePipe
    
    rv5 = CBWTest_New(0, rv4, ChipString)
    
    If rv5 = 1 Then
        rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
        If rv5 <> 1 Then
            rv5 = 2
            Tester.Print "MS bus width Fail"
        End If
    End If
                
    'If rv5 = 1 Then
    '    rv5 = CBWTest_New_128_Sector_AU6377(0, rv5)
    'End If
    ClosePipe
    Call LabelMenu(31, rv5, rv4)
    Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    
                 
    CardResult = DO_WritePort(card, Channel_P1A, &H7F)  'Disconnect all card check NBMD
    Call MsecDelay(0.2)
                
    If GetDeviceName(ChipString) <> "" Then
        rv0 = 3
    End If
             
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
    
    If (rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS) And (Not ((CurV33 = "3") And (CurV18 = "1.6"))) Then
        V18_Count = V18_Count + 1
        If CurV18 = "1.6" Then
            V33_Count = V33_Count + 1
            V18_Count = 0
        End If
        GoTo Routine_Loop
    End If
                
                
                
AU6371ELResult:
    Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)

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

Public Sub AU6433BLF0ETestSub()

'2011/6/14 for FT3 test process (CH1:V33: 3.6V/3.0V ; CH2: V18: 2.2V/1.75V)
'2011/6/13 Just modify O/S pattern (pin2:open,pin11:short)
'201012/30 This code just for V5 S/B
'reduce NBMD test time
'2011/3/24 modify LV: 1.8 -> 1.75 , add MSpro R/W 128 Sector
'2011/3/30 modify MPpro & XD R/W 2K => 4K

Dim TmpLBA As Long
Dim i As Integer
Dim HV_Flag As Boolean
Dim HV_Result As String
Dim LV_Result As String
              
If PCI7248InitFinish = 0 Then
    PCI7248ExistAU6254
    Call SetTimer_1ms
End If
              
'CardResult = DO_WritePort(card, Channel_P1C, &HFF)
'Call PowerSet2(1, "0.0", "0.05", 1, "0.0", "0.05", 1)
      
  Tester.Print "Begin AU6433DL FT2 Test: "
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim CPRMMODE As Byte
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
               
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
                ChipString = "vid"
                
                '=========================================
                '    POWER on
                '=========================================
Routine_Label:

                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                    If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 Call MsecDelay(0.2)
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                    
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If

                
                '===============================================
                '  SD Card test
                '
               '   Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================

                
                
                If (HV_Flag = False) Then
                    Call PowerSet2(1, "3.6", "0.05", 1, "2.2", "0.05", 1)
                    Tester.Print "Begin HV Test ..."
                Else
                    Call PowerSet2(1, "3.0", "0.05", 1, "1.75", "0.05", 1)
                    Tester.Print vbCrLf & "Begin LV Test ..."
                End If
                
                CardResult = DO_WritePort(card, Channel_P1A, &H7E)  ' set SD card detect down
                Call MsecDelay(0.2)
                
                If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect Down Fail"
                    End
                End If
                
                'Call MsecDelay(1.3)
                rv0 = WaitDevOn(ChipString)
                Call MsecDelay(0.2)
                     
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                
                If CardResult <> 0 Then
                    MsgBox "Read light On fail"
                    End
                End If
                     
                ClosePipe
                     
                
                If rv0 = 1 Then
                    rv0 = CBWTest_New(0, 1, ChipString)
                End If
                      
                      If rv0 = 1 Then
                           rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                           If rv0 <> 1 Then
                           rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                       End If
                      
                       If rv0 <> 1 Then ' for E55 command
                          rv0 = Read_SD_SpeedE55(0, 0, 64, "8Bits")
                          If rv0 <> 1 Then
                          rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                          If rv0 = 1 Then
                            CPRMMODE = 1
                          End If
                      End If
                      ClosePipe
                      
                      If (rv0 = 0) Or (rv0 = 2) Then 'for OSE request
                        rv0 = 3
                        GoTo AU6371ELResult
                      End If
                      
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                          Tester.Print "LightON="; LightOn
                          Tester.Print "LightOFF="; LightOff
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                    
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      If rv0 = 1 Then
                        TmpLBA = LBA
                        'LBA = 99
                         'For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6371ELResult
                             End If
                         'Next
                        LBA = TmpLBA
                      End If
'=======================================================================================
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.02)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
                  Call MsecDelay(0.02)
                 
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
                 Call MsecDelay(0.02)
               
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
                 Call MsecDelay(0.02)
                  OpenPipe
                  rv5 = ReInitial(0)
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                
                If CPRMMODE = 0 Then  ' for E54 Before
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                 Else             ' for AU6433E55 after
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476E55(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                End If
                
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H7F)  'Disconnect all card check NBMD
                Call MsecDelay(0.2)
                
                If GetDeviceName(ChipString) <> "" Then
                    rv0 = 3
                End If
                
                
                
                
AU6371ELResult:
                'Call PowerSet2(1, "0.0", "0.05", 1, "0.0", "0.05", 1)
                'CardResult = DO_WritePort(card, Channel_P1A, &H80)
                
                        
                If HV_Flag = False Then
                    If rv0 * rv1 * rv2 * rv3 * rv4 * rv5 = 0 Then
                        HV_Result = "Bin2"
                        Tester.Print "HV Unknow"
                    ElseIf rv0 * rv1 * rv2 * rv3 * rv4 * rv5 <> 1 Then
                        HV_Result = "Fail"
                        Tester.Print "HV Fail"
                    ElseIf rv0 * rv1 * rv2 * rv3 * rv4 * rv5 = 1 Then
                        HV_Result = "PASS"
                        Tester.Print "HV PASS"
                    End If
                    rv0 = 0
                    rv1 = 0
                    rv2 = 0
                    rv3 = 0
                    rv4 = 0
                    rv5 = 0
                    HV_Flag = True
                    Call MsecDelay(0.2)
                    GoTo Routine_Label
                Else
                    If rv0 * rv1 * rv2 * rv3 * rv4 * rv5 = 0 Then
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
                
End Sub

Public Sub AU6433BLF0FTestSub()

'2011/6/14 for FT3 test process (CH1:V33: 3.6V/3.0V ; CH2: V18: 2.2V/1.75V)
'2011/6/13 Just modify O/S pattern (pin2:open,pin11:short)
'201012/30 This code just for V5 S/B
'reduce NBMD test time
'2011/3/24 modify LV: 1.8 -> 1.75 , add MSpro R/W 128 Sector
'2011/3/30 modify MPpro & XD R/W 2K => 4K
'2011/12/6 Add MSpro R/W 64K

Dim TmpLBA As Long
Dim i As Integer
Dim HV_Flag As Boolean
Dim HV_Result As String
Dim LV_Result As String
              
If PCI7248InitFinish = 0 Then
    PCI7248ExistAU6254
    Call SetTimer_1ms
End If
              
'CardResult = DO_WritePort(card, Channel_P1C, &HFF)
'Call PowerSet2(1, "0.0", "0.05", 1, "0.0", "0.05", 1)
      
  Tester.Print "Begin AU6433DL FT3 Test: "
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim CPRMMODE As Byte
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
               
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
                ChipString = "vid"
                
                '=========================================
                '    POWER on
                '=========================================
Routine_Label:

                
                CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                End If
                Call MsecDelay(0.2)
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                    
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If

                
                '===============================================
                '  SD Card test
                '
               '   Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================

                
                
                If (HV_Flag = False) Then
                    Call PowerSet2(1, "3.6", "0.06", 1, "2.2", "0.06", 1)
                    Tester.Print "Begin HV Test ..."
                Else
                    Call PowerSet2(1, "3.0", "0.06", 1, "1.75", "0.06", 1)
                    Tester.Print vbCrLf & "Begin LV Test ..."
                End If
                Call MsecDelay(0.2)
                CardResult = DO_WritePort(card, Channel_P1A, &H7E)  ' set SD card detect down
                Call MsecDelay(0.1)
                
                If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect Down Fail"
                    End
                End If
                
                'Call MsecDelay(1.3)
                rv0 = WaitDevOn(ChipString)
                Call MsecDelay(0.2)
                     
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                
                If CardResult <> 0 Then
                    MsgBox "Read light On fail"
                    End
                End If
                     
                ClosePipe
                     
                
                If rv0 = 1 Then
                    rv0 = CBWTest_New(0, 1, ChipString)
                End If
                      
                If rv0 = 1 Then
                    rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
                    End If
                End If
                      
                If rv0 <> 1 Then ' for E55 command
                    rv0 = Read_SD_SpeedE55(0, 0, 64, "8Bits")
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
                    End If
                    
                    If rv0 = 1 Then
                        CPRMMODE = 1
                    End If
                End If
                    
                      
                If rv0 <> 0 Then
                    If LightOn <> &HBF Or LightOff <> &HFF Then
                        Tester.Print "LightON="; LightOn
                        Tester.Print "LightOFF="; LightOff
                        UsbSpeedTestResult = GPO_FAIL
                        rv0 = 3
                    End If
                End If
                ClosePipe
                
'=======================================================================================
    'SD R / W
'=======================================================================================
                If rv0 = 1 Then
                    TmpLBA = LBA
                    LBA = LBA + 199
                            
                    ClosePipe
                    rv0 = CBWTest_New_128_Sector_AU6377(0, rv0)  ' write
                    ClosePipe
                    LBA = TmpLBA
                End If
'=======================================================================================
                     
                Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  CF Card test
                '================================================
               
                rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                ' Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                
                CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                End If
                  
                  
                Call MsecDelay(0.05)
                CardResult = DO_WritePort(card, Channel_P1A, &H77)
                Call MsecDelay(0.05)
                 
                OpenPipe
                rv3 = ReInitial(0)
                ClosePipe
                
                If rv3 = 1 Then
                    rv3 = CBWTest_New(0, rv2, ChipString)
                    ClosePipe
                End If
                Call LabelMenu(2, rv3, rv2)
                 
                Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  MS Card test
                '================================================
                   
                     
                rv4 = rv3  'AU6344 has no MS slot pin
               
                'Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  MS Pro Card test
                '================================================
              
                CardResult = DO_WritePort(card, Channel_P1A, &H57)
                Call MsecDelay(0.05)
                CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                Call MsecDelay(0.05)
                
                OpenPipe
                rv5 = ReInitial(0)
                ClosePipe
                
                If rv5 = 1 Then
                    rv5 = CBWTest_New(0, rv4, ChipString)
                End If
                
                If CPRMMODE = 0 Then  ' for E54 Before
                    If rv5 = 1 Then
                        rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                        If rv5 <> 1 Then
                            rv5 = 2
                            Tester.Print "MS bus width Fail"
                        End If
                    End If
                 Else             ' for AU6433E55 after
                    If rv5 = 1 Then
                        rv5 = Read_MS_Speed_AU6476E55(0, 0, 64, "4Bits")
                        If rv5 <> 1 Then
                            rv5 = 2
                            Tester.Print "MS bus width Fail"
                        End If
                    End If
                End If
                
                ClosePipe
                
                If rv5 = 1 Then
                    rv5 = CBWTest_New_128_Sector_AU6377(0, rv5)
                    ClosePipe
                End If
                
                Call LabelMenu(31, rv5, rv4)
                Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H7F)  'Disconnect all card check NBMD
                Call MsecDelay(0.2)
                
                If GetDeviceName(ChipString) <> "" Then
                    rv0 = 3
                End If
                
                CardResult = DO_WritePort(card, Channel_P1A, &HFF)  'Disconnect all card check NBMD
                'Call MsecDelay(0.2)
                
                
                
AU6371ELResult:

                If HV_Flag = False Then
                    If rv0 * rv1 * rv2 * rv3 * rv4 * rv5 = 0 Then
                        HV_Result = "Bin2"
                        Tester.Print "HV Unknow"
                    ElseIf rv0 * rv1 * rv2 * rv3 * rv4 * rv5 <> 1 Then
                        HV_Result = "Fail"
                        Tester.Print "HV Fail"
                    ElseIf rv0 * rv1 * rv2 * rv3 * rv4 * rv5 = 1 Then
                        HV_Result = "PASS"
                        Tester.Print "HV PASS"
                    End If
                    rv0 = 0
                    rv1 = 0
                    rv2 = 0
                    rv3 = 0
                    rv4 = 0
                    rv5 = 0
                    HV_Flag = True
                    GoTo Routine_Label
                Else
                    If rv0 * rv1 * rv2 * rv3 * rv4 * rv5 = 0 Then
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
                
End Sub


Public Sub AU6433BLS2FTestSub()

'2011/6/24 for AU6433S61-RBL FT2 test (V33,5V-In use external 3.3V, V18 internal)
'2011/6/16 AU6433S61-RBL FT2 test
'2011/6/13 Just modify O/S pattern (pin2:open,pin11:short)
'201012/30 This code just for V5 S/B
'reduce NBMD test time
'2011/3/24 modify LV: 1.8 -> 1.75 , add MSpro R/W 128 Sector
'2011/3/30 modify MPpro & XD R/W 2K => 4K

Dim TmpLBA As Long
Dim i As Integer

If PCI7248InitFinish = 0 Then
    PCI7248ExistAU6254
    Call SetTimer_1ms
End If

Call PowerSet2(1, "3.3", "0.5", 1, "3.3", "0.5", 1)

'OS_Result = 0
'rv0 = 0

'CardResult = DO_WritePort(card, Channel_P1C, &H0)
                 
'MsecDelay (0.3)

'OpenShortTest_Result

'If OS_Result <> 1 Then
'    rv0 = 0                 'OS Fail
'    GoTo AU6371ELResult
'End If

'CardResult = DO_WritePort(card, Channel_P1C, &HFF)
                       
      
  Tester.Print "Begin AU6433DL FT2 Test: "
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim CPRMMODE As Byte
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
               
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
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                    If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 Call MsecDelay(0.2)
                 
                ' CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                ' Call MsecDelay(1#)    'power on time          'NBMD
                ChipString = "vid"
                ' If GetDeviceName(ChipString) <> "" Then
                '    rv0 = 3
                '    GoTo AU6371ELResult
                '
                '  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
               '   Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     'Call MsecDelay(1.3)
                     
                     rv0 = WaitDevOn(ChipString)
                     Call MsecDelay(0.1)
                     
                     
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                      ClosePipe
                      
                      If rv0 = 1 Then
                        rv0 = CBWTest_New(0, 1, ChipString)
                      End If
                      
                      If rv0 = 1 Then
                           rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                           If rv0 <> 1 Then
                           rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                       End If
                      
                       If rv0 <> 1 Then ' for E55 command
                          rv0 = Read_SD_SpeedE55(0, 0, 64, "8Bits")
                          If rv0 <> 1 Then
                          rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                          If rv0 = 1 Then
                            CPRMMODE = 1
                          End If
                      End If
                      ClosePipe
                      
                      If (rv0 = 0) Or (rv0 = 2) Then 'for OSE request
                        rv0 = 3
                        GoTo AU6371ELResult
                      End If
                      
                      'Tester.Print "rv0="; rv0
                     
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                          Tester.Print "LightON="; LightOn
                          Tester.Print "LightOFF="; LightOff
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                    
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      If rv0 = 1 Then
                        TmpLBA = LBA
                        'LBA = 99
                         'For i = 1 To 30
                             'rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv0 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             'If rv0 <> 1 Then
                             '   LBA = TmpLBA
                             '   GoTo AU6371ELResult
                             'End If
                         'Next
                        LBA = TmpLBA
                      End If
'=======================================================================================
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
                 Call MsecDelay(0.1)
               
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
                 Call MsecDelay(0.1)
                  OpenPipe
                  rv5 = ReInitial(0)
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                
                If CPRMMODE = 0 Then  ' for E54 Before
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                 Else             ' for AU6433E55 after
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476E55(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                End If
                
                'If rv5 = 1 Then
                '    rv5 = CBWTest_New_128_Sector_AU6377(0, 1)
                'End If
                
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H7F)  'Disconnect all card check NBMD
                Call MsecDelay(0.2)
                
                If GetDeviceName(ChipString) <> "" Then
                    rv0 = 3
                End If
                
                
                
AU6371ELResult:
                        CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                        Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
                        
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
Public Sub AU6433BLS0FTestSub()

'2011/6/24 for AU6433S61-RBL FT3 (CH1/CH2: 3.6V/3.0V )
'2011/6/14 for FT3 test process (CH1:V33: 3.6V/3.0V ; CH2: V18: 2.2V/1.75V)
'2011/6/13 Just modify O/S pattern (pin2:open,pin11:short)
'201012/30 This code just for V5 S/B
'reduce NBMD test time
'2011/3/24 modify LV: 1.8 -> 1.75 , add MSpro R/W 128 Sector
'2011/3/30 modify MPpro & XD R/W 2K => 4K

Dim TmpLBA As Long
Dim i As Integer
Dim HV_Flag As Boolean
Dim HV_Result As String
Dim LV_Result As String
              
If PCI7248InitFinish = 0 Then
    PCI7248ExistAU6254
    Call SetTimer_1ms
End If
              
'CardResult = DO_WritePort(card, Channel_P1C, &HFF)
'Call PowerSet2(1, "0.0", "0.05", 1, "0.0", "0.05", 1)
      
  Tester.Print "Begin AU6433DL FT2 Test: "
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim CPRMMODE As Byte
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
               
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
                ChipString = "vid"
                
                '=========================================
                '    POWER on
                '=========================================
Routine_Label:

                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                    If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 Call MsecDelay(0.02)
                 
                 '================================================
                   'CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
                   '
                   'If CardResult <> 0 Then
                   ' MsgBox "Read light off fail"
                   ' End
                   'End If

                
                '===============================================
                '  SD Card test
                '
               '   Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================

                
                
                If (HV_Flag = False) Then
                    Call PowerSet2(1, "3.6", "0.2", 1, "3.6", "0.2", 1)
                    Tester.Print "Begin HV Test ..."
                Else
                    Call PowerSet2(1, "3.0", "0.2", 1, "3.0", "0.2", 1)
                    Tester.Print vbCrLf & "Begin LV Test ..."
                End If
                
                CardResult = DO_WritePort(card, Channel_P1A, &H7E)  ' set SD card detect down
                Call MsecDelay(0.2)
                
                If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect Down Fail"
                    End
                End If
                
                'Call MsecDelay(1.3)
                rv0 = WaitDevOn(ChipString)
                Call MsecDelay(0.2)
                     
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                
                If CardResult <> 0 Then
                    MsgBox "Read light On fail"
                    End
                End If
                     
                ClosePipe
                     
                
                If rv0 = 1 Then
                    rv0 = CBWTest_New(0, 1, ChipString)
                End If
                      
                      If rv0 = 1 Then
                           rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                           If rv0 <> 1 Then
                           rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                       End If
                      
                       If rv0 <> 1 Then ' for E55 command
                          rv0 = Read_SD_SpeedE55(0, 0, 64, "8Bits")
                          If rv0 <> 1 Then
                          rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                          If rv0 = 1 Then
                            CPRMMODE = 1
                          End If
                      End If
                      ClosePipe
                      
                      If (rv0 = 0) Or (rv0 = 2) Then 'for OSE request
                        rv0 = 3
                        GoTo AU6371ELResult
                      End If
                      
                        If rv0 <> 0 Then
                          'If LightON <> &HBF Or LightOFF <> &HFF Then
                          If LightOn <> &HBF Then
                          
                          Tester.Print "LightON="; LightOn
                          'Tester.Print "LightOFF="; LightOFF
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                    
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      If rv0 = 1 Then
                        TmpLBA = LBA
                        'LBA = 99
                         'For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6371ELResult
                             End If
                         'Next
                        LBA = TmpLBA
                      End If
'=======================================================================================
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.02)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
                  Call MsecDelay(0.02)
                 
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
                 Call MsecDelay(0.02)
               
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
                 Call MsecDelay(0.02)
                  OpenPipe
                  rv5 = ReInitial(0)
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                
                If CPRMMODE = 0 Then  ' for E54 Before
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                 Else             ' for AU6433E55 after
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476E55(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                End If
                
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H7F)  'Disconnect all card check NBMD
                Call MsecDelay(0.2)
                
                If GetDeviceName(ChipString) <> "" Then
                    rv0 = 3
                End If
                
                
                
                
AU6371ELResult:
                Call PowerSet2(1, "0.0", "0.2", 1, "0.0", "0.2", 1)
                'CardResult = DO_WritePort(card, Channel_P1A, &H80)
                
                        
                If HV_Flag = False Then
                    If rv0 * rv1 * rv2 * rv3 * rv4 * rv5 = 0 Then
                        HV_Result = "Bin2"
                        Tester.Print "HV Unknow"
                    ElseIf rv0 * rv1 * rv2 * rv3 * rv4 * rv5 <> 1 Then
                        HV_Result = "Fail"
                        Tester.Print "HV Fail"
                    ElseIf rv0 * rv1 * rv2 * rv3 * rv4 * rv5 = 1 Then
                        HV_Result = "PASS"
                        Tester.Print "HV PASS"
                    End If
                    rv0 = 0
                    rv1 = 0
                    rv2 = 0
                    rv3 = 0
                    rv4 = 0
                    rv5 = 0
                    HV_Flag = True
                    Call MsecDelay(0.2)
                    GoTo Routine_Label
                Else
                    If rv0 * rv1 * rv2 * rv3 * rv4 * rv5 = 0 Then
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
                
End Sub
Public Sub AU6433BLD3BTestSub()

Dim TmpLBA As Long
Dim i As Integer

               If PCI7248InitFinish = 0 Then
                  PCI7248ExistAU6254
                  Call SetTimer_1ms
               End If

Call PowerSet2(1, "2.2", "0.05", 1, "2.2", "0.05", 1)

OS_Result = 0
rv0 = 0

CardResult = DO_WritePort(card, Channel_P1C, &H0)
                 
MsecDelay (0.2)

OpenShortTest_Result

If OS_Result <> 1 Then
    rv0 = 0                 'OS Fail
    GoTo AU6371ELResult
End If

CardResult = DO_WritePort(card, Channel_P1C, &HFF)
                       
      
  Tester.Print "Begin AU6433EF FT2 Test: NB mode test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim CPRMMODE As Byte
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
               
               ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
               '  CardResult = DO_WritePort(card, Channel_P1B, &H0)
               
                LBA = LBA + 1
                         
                rv0 = 1
                rv1 = 1
                rv2 = 1
                rv3 = 0
                rv4 = 1
                rv5 = 1
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
                 Call MsecDelay(0.1)
                 
                 'CardResult = DO_WritePort(card, Channel_P1A, &H7F)    'skip NBMD
                  
                 'Call MsecDelay(0.2)    'power on time
                'ChipString = "vid"
                
                'If GetDeviceName(ChipString) <> "" Then
                '    rv0 = 0
                '    GoTo AU6371ELResult
                  
                'End If
                 
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
                  Call MsecDelay(0.3)
                  WaitDevOn ("pid")
                  'OpenPipe
                  
                 
                rv3 = CBWTest_New(0, rv2, ChipString)
                ClosePipe
                
                If rv3 = 1 Then
                    
                    Call PowerSet2(1, "1.8", "0.05", 1, "1.8", "0.05", 1)
                    Call MsecDelay(0.2)
                    rv3 = CBWTest_New(0, rv2, ChipString)
        
                End If
                
                Call LabelMenu(2, rv3, rv2)
                 
                Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                 
                 '================================================
                     
                     
                      
                      
                     
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
Public Sub AU6433BLF25TestSub()

Call PowerSet2(1, "3.3", "0.05", 1, "3.3", "0.05", 1)
      
  Tester.Print "AU6433EF : NB mode test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim CPRMMODE As Byte
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
                ChipString = "vid"
                 If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
               '   Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1.3)
                     
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
                      
                       If rv0 <> 1 Then ' for E55 command
                          rv0 = Read_SD_SpeedE55(0, 0, 64, "8Bits")
                          If rv0 <> 1 Then
                          rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                          If rv0 = 1 Then
                            CPRMMODE = 1
                          End If
                      End If
                      ClosePipe
                      
                      Tester.Print "rv0="; rv0
                     
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                          Tester.Print "LightON="; LightOn
                          Tester.Print "LightOFF="; LightOff
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
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
                 Call MsecDelay(0.1)
               
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
                 Call MsecDelay(0.1)
                  OpenPipe
                  rv5 = ReInitial(0)
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                
                If CPRMMODE = 0 Then  ' for E54 Before
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                 Else             ' for AU6433E55 after
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476E55(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
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

Public Sub AU6433BLF27TestSub()
Dim GPIPin As Byte
Call PowerSet2(0, "3.3", "0.05", 1, "3.3", "0.05", 1)
      
  Tester.Print "AU6433EF : NB mode test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim CPRMMODE As Byte
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
                ChipString = "vid_1984"
                   If GetDeviceName(ChipString) <> "" Then
                      rv0 = 0
                      GoTo AU6371ELResult
                 
                   End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
                 
               
                '===============================================
                '  SD Card test
                '
               '   Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1.3)
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
                      
                       If rv0 <> 1 Then ' for E55 command
                          rv0 = Read_SD_SpeedE55(0, 0, 64, "8Bits")
                          If rv0 <> 1 Then
                          rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                          If rv0 = 1 Then
                            CPRMMODE = 1
                          End If
                      End If
                      
                       GPIPin = Read_GPI(0, 0, 64)
                      
                      ClosePipe
                       Tester.Print "Port2:"; Hex(LightOn)
                       Tester.Print "GPIPin:"; Hex(GPIPin)
                      
                       If LightOn <> &HBC Or GPIPin <> &H15 Then
                        Tester.Print "OS Fail"
                        rv0 = 3
                         Tester.Label9.Caption = "OS Test Fail"
                          GoTo AU6371ELResult
                      End If
                       
                      
                      Tester.Print "rv0="; rv0
                     
                        If rv0 <> 0 Then
                        '  If LightON <> &HBF Or LightOFF <> &HFF Then
                        '  Tester.Print "LightON="; LightON
                        '  Tester.Print "LightOFF="; LightOFF
                        '  UsbSpeedTestResult = GPO_FAIL
                        '  rv0 = 3
                        '  End If
                        End If
                    
                     
                     
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
                 Call MsecDelay(0.1)
               
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
                 Call MsecDelay(0.1)
                  OpenPipe
                  rv5 = ReInitial(0)
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                
                If CPRMMODE = 0 Then  ' for E54 Before
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                 Else             ' for AU6433E55 after
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476E55(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
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
Public Sub AU6433DLF20TestSub()

Dim TmpLBA As Long
Dim i As Integer
Call PowerSet2(1, "3.3", "0.05", 1, "2.1", "0.05", 1)
      
  Tester.Print "AU6433EF : NB mode test , change 1.8V to 2.1V"
'==================================================================
'
'  this code come from AU6433BLF29ELTestSub
'  just LightON form &HBF to &H7F
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim CPRMMODE As Byte
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
                ChipString = "vid"
                 If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
               '   Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1.3)
                     
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
                      
                       If rv0 <> 1 Then ' for E55 command
                          rv0 = Read_SD_SpeedE55(0, 0, 64, "8Bits")
                          If rv0 <> 1 Then
                          rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                          If rv0 = 1 Then
                            CPRMMODE = 1
                          End If
                      End If
                      ClosePipe
                      
                      Tester.Print "rv0="; rv0
                     
                        If rv0 <> 0 Then
                          If LightOn <> &H7F Or LightOff <> &HFF Then
                          Tester.Print "LightON="; LightOn
                          Tester.Print "LightOFF="; LightOff
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                    
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      
                     TmpLBA = LBA
                     'LBA = 99
                         'For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6371ELResult
                             End If
                         'Next
                      LBA = TmpLBA
                      
'=======================================================================================
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
                 Call MsecDelay(0.1)
               
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
                 Call MsecDelay(0.1)
                  OpenPipe
                  rv5 = ReInitial(0)
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                
                If CPRMMODE = 0 Then  ' for E54 Before
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                 Else             ' for AU6433E55 after
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476E55(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
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
Public Sub AU6433DLF21TestSub()

Dim TmpLBA As Long
Dim i As Integer
      
  Tester.Print "AU6433DL Using Internal LDO Output..."
'==================================================================
'
'  this code come from AU6433BLF29ELTestSub
'  just LightON form &HBF to &H7F
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim CPRMMODE As Byte
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.2)
                                  
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 'Call MsecDelay(1#)    'power on time
                ChipString = "vid"
                 'If GetDeviceName(ChipString) <> "" Then
                 '   rv0 = 0
                 '   GoTo AU6371ELResult
                  
                 ' End If
                 
                 '================================================
                   
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
               '   Call MsecDelay(0.3)
             
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
                    rv0 = WaitDevOn(ChipString)
                    Call MsecDelay(0.1)
                     
                    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      
                    If CardResult <> 0 Then
                        MsgBox "Read light On fail"
                        End
                    End If
                           
                    ClosePipe
                      
                    If rv0 = 1 Then
                        rv0 = CBWTest_New(0, 1, ChipString)
                    End If
                    
                    If rv0 = 1 Then
                        rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                        If rv0 <> 1 Then
                            rv0 = 2
                            Tester.Print "SD bus width Fail"
                        End If
                    End If
                      
                    If rv0 <> 1 Then ' for E55 command
                        rv0 = Read_SD_SpeedE55(0, 0, 64, "8Bits")
                        If rv0 <> 1 Then
                            rv0 = 2
                            Tester.Print "SD bus width Fail"
                        End If
                        
                        If rv0 = 1 Then
                            CPRMMODE = 1
                        End If
                    End If
                    
                    ClosePipe
                    
                    Tester.Print "rv0="; rv0
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      
                     TmpLBA = LBA
                     'LBA = 99
                         'For i = 1 To 30
                        rv1 = 0
                        LBA = LBA + 199
                            
                        ClosePipe
                        
                        rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             
                        If rv1 <> 1 Then
                            LBA = TmpLBA
                            GoTo AU6371ELResult
                        End If
                         'Next
                        LBA = TmpLBA
                      
'=======================================================================================
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  'rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                
                CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                End If
                  
                  
                Call MsecDelay(0.1)
                
                CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                   
                     
                rv4 = rv3  'AU6344 has no MS slot pin
               
                '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
                Call MsecDelay(0.1)
               
                CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
                Call MsecDelay(0.1)
                
                OpenPipe
                rv5 = ReInitial(0)
                ClosePipe
                
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                
                If CPRMMODE = 0 Then  ' for E54 Before
                    If rv5 = 1 Then
                        rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                        
                        If rv5 <> 1 Then
                            rv5 = 2
                            Tester.Print "MS bus width Fail"
                        End If
                    End If
                Else             ' for AU6433E55 after
                
                    If rv5 = 1 Then
                        rv5 = Read_MS_Speed_AU6476E55(0, 0, 64, "4Bits")
                        If rv5 <> 1 Then
                            rv5 = 2
                            Tester.Print "MS bus width Fail"
                        End If
                    End If
                End If
                
                Call LabelMenu(31, rv5, rv4)
                
                Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                ClosePipe
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H7F)   'Check NB mode
                Call MsecDelay(0.2)
                
                If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    Tester.Print "NB Mode test fail!"
                    GoTo AU6371ELResult
                End If
                
                CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                    
                Tester.Print "rv0="; rv0
                     
                If rv0 <> 0 Then
                    If LightOn <> &H7F Or LightOff <> &HFF Then
                        Tester.Print "LightON="; LightOn
                        Tester.Print "LightOFF="; LightOff
                        UsbSpeedTestResult = GPO_FAIL
                        rv0 = 3
                    End If
                End If
                
                Call LabelMenu(31, rv5, rv0)
                  
AU6371ELResult:

                CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
                
                
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
Public Sub AU6433DLF22TestSub()

Dim TmpLBA As Long
Dim i As Integer
      
  Tester.Print "AU6433DL Using Internal LDO Output..."
'==================================================================
'
'  this code come from AU6433BLF29ELTestSub
'  just LightON form &HBF to &H7F
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim CPRMMODE As Byte
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.2)
                                  
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 'Call MsecDelay(1#)    'power on time
                ChipString = "vid"
                 'If GetDeviceName(ChipString) <> "" Then
                 '   rv0 = 0
                 '   GoTo AU6371ELResult
                  
                 ' End If
                 
                 '================================================
                   
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
               '   Call MsecDelay(0.3)
             
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
                    rv0 = WaitDevOn(ChipString)
                    Call MsecDelay(0.1)
                       
                       
                    If CardResult <> 0 Then
                        MsgBox "Read light On fail"
                        End
                    End If
                           
                    ClosePipe
                      
                    If rv0 = 1 Then
                        CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                        Call MsecDelay(0.02)
                        
                        If LightOn <> 127 Then
                            Call MsecDelay(0.1)
                            CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                            Call MsecDelay(0.02)
                            
                            If LightOn <> 127 Then
                                Call MsecDelay(0.1)
                                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                            End If
                        End If
                        
                        rv0 = CBWTest_New(0, 1, ChipString)
                    End If
                    
                    
                    
                    If rv0 = 1 Then
                        rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                        If rv0 <> 1 Then
                            rv0 = 2
                            Tester.Print "SD bus width Fail"
                        End If
                    End If
                      
                    If rv0 <> 1 Then ' for E55 command
                        rv0 = Read_SD_SpeedE55(0, 0, 64, "8Bits")
                        If rv0 <> 1 Then
                            rv0 = 2
                            Tester.Print "SD bus width Fail"
                        End If
                        
                        If rv0 = 1 Then
                            CPRMMODE = 1
                        End If
                    End If
                    
                    ClosePipe
                    
                    Tester.Print "rv0="; rv0
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      
                     TmpLBA = LBA
                     'LBA = 99
                         'For i = 1 To 30
                        rv1 = 0
                        LBA = LBA + 199
                            
                        ClosePipe
                        
                        rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             
                        If rv1 <> 1 Then
                            LBA = TmpLBA
                            GoTo AU6371ELResult
                        End If
                         'Next
                        LBA = TmpLBA
                      
'=======================================================================================
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  'rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                
                CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                End If
                  
                  
                Call MsecDelay(0.1)
                
                CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                   
                     
                rv4 = rv3  'AU6344 has no MS slot pin
               
                '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
                Call MsecDelay(0.1)
               
                CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
                Call MsecDelay(0.1)
                
                OpenPipe
                rv5 = ReInitial(0)
                ClosePipe
                
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                
                If CPRMMODE = 0 Then  ' for E54 Before
                    If rv5 = 1 Then
                        rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                        
                        If rv5 <> 1 Then
                            rv5 = 2
                            Tester.Print "MS bus width Fail"
                        End If
                    End If
                Else             ' for AU6433E55 after
                
                    If rv5 = 1 Then
                        rv5 = Read_MS_Speed_AU6476E55(0, 0, 64, "4Bits")
                        If rv5 <> 1 Then
                            rv5 = 2
                            Tester.Print "MS bus width Fail"
                        End If
                    End If
                End If
                
                Call LabelMenu(31, rv5, rv4)
                
                Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                ClosePipe
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H7F)   'Check NB mode
                Call MsecDelay(0.2)
                
                If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    Tester.Print "NB Mode test fail!"
                    GoTo AU6371ELResult
                End If
                
                CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                    
                Tester.Print "rv0="; rv0
                     
                If rv0 <> 0 Then
                    If LightOn <> &H7F Or LightOff <> &HFF Then
                        Tester.Print "LightON="; LightOn
                        Tester.Print "LightOFF="; LightOff
                        UsbSpeedTestResult = GPO_FAIL
                        rv0 = 3
                    End If
                End If
                
                Call LabelMenu(31, rv5, rv0)
                  
AU6371ELResult:

                CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
                
                
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
Public Sub AU6433DLF23TestSub()

Dim TmpLBA As Long
Dim i As Integer
      
  Tester.Print "AU6433DL V18 Using Internal LDO Output..."
'==================================================================
'
'  this code come from AU6433BLF29ELTestSub
'  just LightON form &HBF to &H7F
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim CPRMMODE As Byte
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                    If CardResult <> 0 Then
                        MsgBox "Power off fail"
                    End
                End If
                Call MsecDelay(0.2)
                                  
                CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                
                If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                End If
                   
                     
                     
                ' CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 
                ChipString = "vid_058f"
                 
                
                '===============================================
                '  SD Card test
                '
               '   Call MsecDelay(0.3)
             
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
                    rv0 = WaitDevOn(ChipString)
                    Call MsecDelay(0.1)
                       
                    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                    
                    If CardResult <> 0 Then
                        MsgBox "Read light On fail"
                        End
                    End If
                    
                    Call MsecDelay(0.02)
                           
                    If rv0 <> 0 Then
                        If LightOn <> &H1F Or LightOff <> &H9F Then
                        'If (CAndValue(LightON, &H1) <> 0) Or (CAndValue(LightOFF, &H1) <> 1) Then
                            Tester.Print "LightON="; LightOn
                            Tester.Print "LightOFF="; LightOff
                            UsbSpeedTestResult = GPO_FAIL
                            rv0 = 3
                        End If
                    End If
                           
                    ClosePipe
                    
                    If rv0 = 1 Then
                        rv0 = CBWTest_New(0, 1, ChipString)
                    End If
                    
                    If rv0 = 1 Then
                        rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                        If rv0 <> 1 Then
                            rv0 = 2
                            Tester.Print "SD bus width Fail"
                        End If
                    End If
                      
                    If rv0 <> 1 Then ' for E55 command
                        rv0 = Read_SD_SpeedE55(0, 0, 64, "8Bits")
                        If rv0 <> 1 Then
                            rv0 = 2
                            Tester.Print "SD bus width Fail"
                        End If
                        
                        If rv0 = 1 Then
                            CPRMMODE = 1
                        End If
                    End If
                    
                    ClosePipe
                    
'=======================================================================================
    'SD R / W
'=======================================================================================
                      
                    If rv0 = 1 Then
                        TmpLBA = LBA
                        LBA = LBA + 199
                            
                        rv0 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                        LBA = TmpLBA
                    End If
'=======================================================================================
                     
                Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                
                CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                End If
                  
                  
                Call MsecDelay(0.1)
                
                CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                   
                     
                rv4 = rv3  'AU6344 has no MS slot pin
               
                '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  MS Pro Card test
                '================================================
              
                CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
                Call MsecDelay(0.1)
               
                CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
                Call MsecDelay(0.1)
                
                OpenPipe
                rv5 = ReInitial(0)
                ClosePipe
                
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                
                If CPRMMODE = 0 Then  ' for E54 Before
                    If rv5 = 1 Then
                        rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                        
                        If rv5 <> 1 Then
                            rv5 = 2
                            Tester.Print "MS bus width Fail"
                        End If
                    End If
                Else             ' for AU6433E55 after
                
                    If rv5 = 1 Then
                        rv5 = Read_MS_Speed_AU6476E55(0, 0, 64, "4Bits")
                        If rv5 <> 1 Then
                            rv5 = 2
                            Tester.Print "MS bus width Fail"
                        End If
                    End If
                End If
                
                Call LabelMenu(31, rv5, rv4)
                
                Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                ClosePipe
                 
                If rv5 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H7F)   'Check NB mode
                    Call MsecDelay(0.2)
                
                    If GetDeviceName(ChipString) <> "" Then
                        rv0 = 0
                        Tester.Print "NB Mode test fail!"
                    End If
                
                    Call LabelMenu(31, rv5, rv0)
                End If
                  
AU6371ELResult:

                CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
                
                
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


Public Sub AU6433DLF03TestSub()

Dim TmpLBA As Long
Dim i As Integer
      
Tester.Print "Begin AU6433DL HV+LV Test..."

'==================================================================
'
'  this code come from AU6433BLF29ELTestSub
'  just LightON form &HBF to &H7F
'  2011/8/19: for CSMC HV+LV FT test
'
'===================================================================


    Dim ChipString As String
    Dim AU6371EL_SD As Byte
    Dim AU6371EL_CF As Byte
    Dim AU6371EL_XD As Byte
    Dim AU6371EL_MS As Byte
    Dim AU6371EL_MSP  As Byte
    Dim AU6371EL_BootTime As Single
    Dim HV_Flag As Boolean
    Dim HV_Result As String
    Dim LV_Result As String
    Dim CPRMMODE As Byte
                

        'initial condition
        
        OldChipName = ""
        AU6371EL_SD = 1
        AU6371EL_CF = 2
        AU6371EL_XD = 8
        AU6371EL_MS = 32
        AU6371EL_MSP = 64
        AU6371EL_BootTime = 0.6
                    
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
        
        CardResult = DO_WritePort(card, Channel_P1A, &H7F)
        
        If CardResult <> 0 Then
            MsgBox "Power off fail"
            End
        End If
        
        Call MsecDelay(0.2)
                                  
        CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                
        If CardResult <> 0 Then
            MsgBox "Read light off fail"
            End
        End If
                 
        ChipString = "vid_058f"
                 
                
    '===============================================
    '  SD Card test
    '===============================================
    
    ' set SD card detect down
        

Routine_Label_AU6433DLF03:


        If (HV_Flag = False) Then
            Call PowerSet2(1, "3.6", "0.5", 1, "3.6", "0.5", 1)
            Tester.Print "Begin HV Test ..."
        Else
            Call PowerSet2(1, "3.1", "0.5", 1, "3.1", "0.5", 1)
            Tester.Print vbCrLf & "Begin LV Test ..."
            Call MsecDelay(0.2)
        End If
        
        
        CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
        If CardResult <> 0 Then
            MsgBox "Set SD Card Detect Down Fail"
            End
        End If
                     
        'Call MsecDelay(0.2)
        rv0 = WaitDevOn(ChipString)
        Call MsecDelay(0.1)
                       
        CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                    
        If CardResult <> 0 Then
            MsgBox "Read light On fail"
            End
        End If
                    
        Call MsecDelay(0.02)
                           
        If rv0 <> 0 Then
            'If LightON <> &H1F Or LightOFF <> &H9F Then
            If (CAndValue(LightOn, &H80) <> 0) Then
                Tester.Print "LightON="; LightOn
                Tester.Print "LightOFF="; LightOff
                UsbSpeedTestResult = GPO_FAIL
                rv0 = 3
            End If
        End If
                           
        ClosePipe
                    
        If rv0 = 1 Then
            rv0 = CBWTest_New(0, 1, ChipString)
        End If
                    
        If rv0 = 1 Then
            rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
            
            If rv0 <> 1 Then
                rv0 = 2
                Tester.Print "SD bus width Fail"
            End If
        End If
                      
        If rv0 <> 1 Then ' for E55 command
            rv0 = Read_SD_SpeedE55(0, 0, 64, "8Bits")
            If rv0 <> 1 Then
                rv0 = 2
                Tester.Print "SD bus width Fail"
            End If
                        
            If rv0 = 1 Then
                CPRMMODE = 1
            End If
        End If
                    
        ClosePipe
                    
        '=======================================================================================
        '   SD R / W
        '=======================================================================================
                      
        If rv0 = 1 Then
            TmpLBA = LBA
            LBA = LBA + 199
                            
            rv0 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
            LBA = TmpLBA
        End If
                     
        Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
        Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                
    '===============================================
    '  CF Card test
    '================================================
               
        rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
        ' Call LabelMenu(1, rv1, rv0)
            
        ' Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    '===============================================
    '  SMC Card test  : stop these test for card not enough
    '================================================
              
        rv2 = rv1   ' 'AU6433 has no SMC slot
               
              
    '===============================================
    '  XD Card test
    '================================================
                
        CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
        If CardResult <> 0 Then
            MsgBox "Set XD Card Detect On Fail"
            End
        End If
                  
        Call MsecDelay(0.1)
        CardResult = DO_WritePort(card, Channel_P1A, &H77)
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
                     
        rv4 = rv3  'AU6344 has no MS slot pin
               
        'Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    
    '===============================================
    '  MS Pro Card test
    '================================================
              
        CardResult = DO_WritePort(card, Channel_P1A, &H57)
        Call MsecDelay(0.1)
        CardResult = DO_WritePort(card, Channel_P1A, &H5F)
        Call MsecDelay(0.1)
                
        OpenPipe
        rv5 = ReInitial(0)
        ClosePipe
                
        rv5 = CBWTest_New(0, rv4, ChipString)
                
        If CPRMMODE = 0 Then  ' for E54 Before
            If rv5 = 1 Then
                rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                        
                If rv5 <> 1 Then
                    rv5 = 2
                    Tester.Print "MS bus width Fail"
                End If
            End If
        Else             ' for AU6433E55 after
                
            If rv5 = 1 Then
                rv5 = Read_MS_Speed_AU6476E55(0, 0, 64, "4Bits")
                
                If rv5 <> 1 Then
                    rv5 = 2
                    Tester.Print "MS bus width Fail"
                End If
            End If
        End If
                
        Call LabelMenu(31, rv5, rv4)
                
        Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
        ClosePipe
                 
        If rv5 = 1 Then
            CardResult = DO_WritePort(card, Channel_P1A, &H7F)   'Check NB mode
            Call MsecDelay(0.2)
                
            If GetDeviceName(ChipString) <> "" Then
                rv0 = 0
                Tester.Print "NB Mode test fail!"
            End If
                
            Call LabelMenu(31, rv5, rv0)
        End If
                  
                  
                  
AU6371ELResult:

        
        CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
        Call PowerSet2(1, "0.0", "0.2", 1, "0.0", "0.2", 1)
    
        If HV_Flag = False Then
            If rv0 * rv1 * rv2 * rv3 * rv4 * rv5 = 0 Then
                HV_Result = "Bin2"
                Tester.Print "HV Unknow"
            ElseIf rv0 * rv1 * rv2 * rv3 * rv4 * rv5 <> 1 Then
                HV_Result = "Fail"
                Tester.Print "HV Fail"
            ElseIf rv0 * rv1 * rv2 * rv3 * rv4 * rv5 = 1 Then
                HV_Result = "PASS"
                Tester.Print "HV PASS"
            End If
            
            ReaderExist = 0
            HV_Flag = True
            Call MsecDelay(0.2)
            GoTo Routine_Label_AU6433DLF03
        Else
            If rv0 * rv1 * rv2 * rv3 * rv4 * rv5 = 0 Then
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
        
End Sub
Public Sub AU6433DLF30TestSub()

Dim TmpLBA As Long
Dim i As Integer


               If PCI7248InitFinish = 0 Then
                  PCI7248ExistAU6254
                  Call SetTimer_1ms
               End If

Call PowerSet2(1, "3.3", "0.05", 1, "2.1", "0.05", 1)

OS_Result = 0
rv0 = 0

CardResult = DO_WritePort(card, Channel_P1C, &H0)
                 
MsecDelay (0.3)

OpenShortTest_Result

If OS_Result <> 1 Then
    rv0 = 0                 'OS Fail
    GoTo AU6371ELResult
End If

CardResult = DO_WritePort(card, Channel_P1C, &HFF)
      
      
      
  Tester.Print "AU6433EF : NB mode test , change 1.8V to 2.1V"
'==================================================================
'
'  this code come from AU6433BLF29ELTestSub
'  just LightON form &HBF to &H7F
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim CPRMMODE As Byte
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
                ChipString = "vid"
                 If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
               '   Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1.3)
                     
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
                      
                       If rv0 <> 1 Then ' for E55 command
                          rv0 = Read_SD_SpeedE55(0, 0, 64, "8Bits")
                          If rv0 <> 1 Then
                          rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                          If rv0 = 1 Then
                            CPRMMODE = 1
                          End If
                      End If
                      ClosePipe
                      
                      Tester.Print "rv0="; rv0
                     
                        If rv0 <> 0 Then
                          If LightOn <> &H7F Or LightOff <> &HFF Then
                          Tester.Print "LightON="; LightOn
                          Tester.Print "LightOFF="; LightOff
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                    
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      
                     TmpLBA = LBA
                     'LBA = 99
                         'For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6371ELResult
                             End If
                         'Next
                      LBA = TmpLBA
                      
'=======================================================================================
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
                 Call MsecDelay(0.1)
               
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
                 Call MsecDelay(0.1)
                  OpenPipe
                  rv5 = ReInitial(0)
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                
                If CPRMMODE = 0 Then  ' for E54 Before
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                 Else             ' for AU6433E55 after
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476E55(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
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
Public Sub AU6433DLF31TestSub()


'201012/30 This code from AU6433BLF3B
'just LightON form &HBF to &H7F

Dim TmpLBA As Long
Dim i As Integer

               If PCI7248InitFinish = 0 Then
                  PCI7248ExistAU6254
                  Call SetTimer_1ms
               End If
               
If Dir("D:\LABPC.PC") = "LABPC.PC" Then
    Call PowerSet2(1, "3.3", "0.05", 1, "2.2", "0.05", 1)
Else
    Call PowerSet2(1, "2.2", "0.05", 1, "2.2", "0.05", 1)
End If

OS_Result = 0
rv0 = 0

CardResult = DO_WritePort(card, Channel_P1C, &H0)
                 
MsecDelay (0.3)

OpenShortTest_Result

If OS_Result <> 1 Then
    rv0 = 0                 'OS Fail
    GoTo AU6371ELResult
End If

CardResult = DO_WritePort(card, Channel_P1C, &HFF)
                       
      
  Tester.Print "Begin AU6433EF FT2 Test: NB mode test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim CPRMMODE As Byte
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
               
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
                 Call MsecDelay(0.1)
                 
                ' CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                ' Call MsecDelay(1#)    'power on time          'NBMD
                ChipString = "vid"
                ' If GetDeviceName(ChipString) <> "" Then
                '    rv0 = 3
                '    GoTo AU6371ELResult
                '
                '  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
               '   Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1.3)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      
                      
                      If rv0 = 1 Then
                        If Dir("D:\LABPC.PC") = "LABPC.PC" Then
                            Call PowerSet2(1, "3.3", "0.05", 1, "1.8", "0.05", 1)
                        Else
                            Call PowerSet2(1, "1.8", "0.05", 1, "1.8", "0.05", 1)
                        End If
                        
                        rv0 = CBWTest_New(0, 1, ChipString)
                      End If
                      
                      If rv0 = 1 Then
                           rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                           If rv0 <> 1 Then
                           rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                       End If
                      
                       If rv0 <> 1 Then ' for E55 command
                          rv0 = Read_SD_SpeedE55(0, 0, 64, "8Bits")
                          If rv0 <> 1 Then
                          rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                          If rv0 = 1 Then
                            CPRMMODE = 1
                          End If
                      End If
                      ClosePipe
                      
                      If (rv0 = 0) Or (rv0 = 2) Then 'for OSE request
                        rv0 = 2
                        GoTo AU6371ELResult
                      End If
                      
                      Tester.Print "rv0="; rv0
                     
                        If rv0 <> 0 Then
                          If LightOn <> &H7F Or LightOff <> &HFF Then
                          Tester.Print "LightON="; LightOn
                          Tester.Print "LightOFF="; LightOff
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                    
'=======================================================================================
    'SD R / W
'=======================================================================================
                      If rv0 = 1 Then
                        TmpLBA = LBA
                        'LBA = 99
                         'For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6371ELResult
                             End If
                         'Next
                        LBA = TmpLBA
                      End If
'=======================================================================================
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
                 Call MsecDelay(0.1)
               
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
                 Call MsecDelay(0.1)
                  OpenPipe
                  rv5 = ReInitial(0)
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                
                If CPRMMODE = 0 Then  ' for E54 Before
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                 Else             ' for AU6433E55 after
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476E55(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                End If
                
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H7F)  'Disconnect all card check NBMD
                Call MsecDelay(0.2)
                
                If GetDeviceName(ChipString) <> "" Then
                    rv0 = 3
                End If
                
                
                
AU6371ELResult:

                        CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                        
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

Public Sub AU6433DLF32TestSub()

Dim TmpLBA As Long
Dim i As Integer

If PCI7248InitFinish = 0 Then
    PCI7248ExistAU6254
    Call SetTimer_1ms
End If

OS_Result = 0
rv0 = 0

CardResult = DO_WritePort(card, Channel_P1C, &H0)
                 
MsecDelay (0.3)

OpenShortTest_Result

If OS_Result <> 1 Then
    rv0 = 0                 'OS Fail
    GoTo AU6371ELResult
End If

CardResult = DO_WritePort(card, Channel_P1C, &HFF)
                       

Tester.Print "AU6433DL Using Internal LDO Output..."

'==================================================================
'
'  this code come from AU6433DLF22TestSub
'  Purpose to solve GPON7 Detect issue
'
'===================================================================


    Dim ChipString As String
                
    Dim AU6371EL_SD As Byte
    Dim AU6371EL_CF As Byte
    Dim AU6371EL_XD As Byte
    Dim AU6371EL_MS As Byte
    Dim AU6371EL_MSP  As Byte
    Dim AU6371EL_BootTime As Single
    Dim CPRMMODE As Byte
    OldChipName = ""
               
                 
                 
    ' initial condition
                
    AU6371EL_SD = 1
    AU6371EL_CF = 2
    AU6371EL_XD = 8
    AU6371EL_MS = 32
    AU6371EL_MSP = 64
                    
    AU6371EL_BootTime = 0.6
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
                  
    ChipString = "vid"
'================================================
                   

'=======================================================================================
    'SD R / W
'=======================================================================================
' set SD card detect down
    CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
    If CardResult <> 0 Then
        MsgBox "Set SD Card Detect Down Fail"
        End
    End If
                     
    Call MsecDelay(0.3)
    rv0 = WaitDevOn(ChipString)
    Call MsecDelay(0.1)
                       
    If CardResult <> 0 Then
        MsgBox "Read light On fail"
        End
    End If
                           
    ClosePipe
                      
    If rv0 = 1 Then
        CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
        Call MsecDelay(0.02)
                        
        If LightOn <> 127 Then
            Call MsecDelay(0.1)
            CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
            Call MsecDelay(0.02)
                            
            If LightOn <> 127 Then
                Call MsecDelay(0.1)
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
            End If
        End If
                        
        rv0 = CBWTest_New(0, 1, ChipString)
    End If
                    
    If rv0 = 1 Then
        rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
        
        If rv0 <> 1 Then
            rv0 = 2
            Tester.Print "SD bus width Fail"
        End If
    End If
                      
    If rv0 <> 1 Then ' for E55 command
        rv0 = Read_SD_SpeedE55(0, 0, 64, "8Bits")
        If rv0 <> 1 Then
            rv0 = 2
            Tester.Print "SD bus width Fail"
        End If
                        
        If rv0 = 1 Then
            CPRMMODE = 1
        End If
    End If
                    
    ClosePipe
                    
    Tester.Print "rv0="; rv0
                     
    If (rv0 = 0) Or (rv0 = 2) Then 'for OSE request
        rv0 = 2
        GoTo AU6371ELResult
    End If
                     
'=======================================================================================
    'SD 128Sector R / W
'=======================================================================================
                      
    TmpLBA = LBA
    rv1 = 0
    LBA = LBA + 199
                            
    ClosePipe
                        
    rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             
    If rv1 <> 1 Then
        LBA = TmpLBA
        GoTo AU6371ELResult
    End If
                      
'=======================================================================================
                     
    Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
    Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                
'===============================================
'  CF Card test
'================================================
               
    'rv1 = rv0  '----------- AU6371S3 dp not have CF slot
    'Call LabelMenu(1, rv1, rv0)
            
    'Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                
'===============================================
'  SMC Card test  : stop these test for card not enough
'================================================
              
    'AU6433 has no SMC slot
               
    rv2 = rv1   ' to complete the SMC asbolish
               
              
'===============================================
'  XD Card test
'================================================
                
    CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
    If CardResult <> 0 Then
        MsgBox "Set XD Card Detect On Fail"
        End
    End If
                  
    Call MsecDelay(0.1)
                
    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                     
    rv4 = rv3  'AU6344 has no MS slot pin
               
    'Call LabelMenu(2, rv4, rv3)
                 
    'Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
               
                
                
'===============================================
'  MS Pro Card test
'================================================
                
    CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
    Call MsecDelay(0.1)
               
    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
    Call MsecDelay(0.1)
                
    OpenPipe
    rv5 = ReInitial(0)
    ClosePipe
                
    rv5 = CBWTest_New(0, rv4, ChipString)
                
                
    If CPRMMODE = 0 Then  ' for E54 Before
        If rv5 = 1 Then
            rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                        
            If rv5 <> 1 Then
                rv5 = 2
                Tester.Print "MS bus width Fail"
            End If
        End If
    Else             ' for AU6433E55 after
                
        If rv5 = 1 Then
            rv5 = Read_MS_Speed_AU6476E55(0, 0, 64, "4Bits")
                If rv5 <> 1 Then
                    rv5 = 2
                    Tester.Print "MS bus width Fail"
                End If
        End If
    End If
                
    Call LabelMenu(31, rv5, rv4)
                
    Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    ClosePipe
                 
    CardResult = DO_WritePort(card, Channel_P1A, &H7F)   'Check NB mode
    Call MsecDelay(0.2)
                
    If GetDeviceName(ChipString) <> "" Then
        rv0 = 0
        Tester.Print "NB Mode test fail!"
        GoTo AU6371ELResult
    End If
                
    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                    
    Tester.Print "rv0="; rv0
                     
    If rv0 <> 0 Then
        If LightOn <> &H7F Or LightOff <> &HFF Then
            Tester.Print "LightON="; LightOn
            Tester.Print "LightOFF="; LightOff
            UsbSpeedTestResult = GPO_FAIL
            rv0 = 3
        End If
    End If
                
    Call LabelMenu(31, rv5, rv0)
                  
                                    
AU6371ELResult:

    CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
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
Public Sub AU6433DLF33TestSub()

'2011/7/13 for GSMC (Just O/S pattern different)
Dim TmpLBA As Long
Dim i As Integer

If PCI7248InitFinish = 0 Then
    PCI7248ExistAU6254
    Call SetTimer_1ms
End If


Call PowerSet2(1, "3.3", "0.2", 1, "3.3", "0.2", 1)

OS_Result = 0
rv0 = 0

CardResult = DO_WritePort(card, Channel_P1C, &H0)
                 
MsecDelay (0.3)

OpenShortTest_Result

If OS_Result <> 1 Then
    rv0 = 0                 'OS Fail
    GoTo AU6371ELResult
End If

CardResult = DO_WritePort(card, Channel_P1C, &HFF)
                       

Tester.Print "AU6433DL Using Internal LDO Output..."

'==================================================================
'
'  this code come from AU6433DLF22TestSub
'  Purpose to solve GPON7 Detect issue
'
'===================================================================


    Dim ChipString As String
                
    Dim AU6371EL_SD As Byte
    Dim AU6371EL_CF As Byte
    Dim AU6371EL_XD As Byte
    Dim AU6371EL_MS As Byte
    Dim AU6371EL_MSP  As Byte
    Dim AU6371EL_BootTime As Single
    Dim CPRMMODE As Byte
    OldChipName = ""
               
                 
                 
    ' initial condition
                
    AU6371EL_SD = 1
    AU6371EL_CF = 2
    AU6371EL_XD = 8
    AU6371EL_MS = 32
    AU6371EL_MSP = 64
                    
    AU6371EL_BootTime = 0.6
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
                  
    ChipString = "vid"
'================================================
                   

'=======================================================================================
    'SD R / W
'=======================================================================================
' set SD card detect down
    CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
    If CardResult <> 0 Then
        MsgBox "Set SD Card Detect Down Fail"
        End
    End If
                     
    Call MsecDelay(0.2)
    rv0 = WaitDevOn(ChipString)
    Call MsecDelay(0.1)
                       
    If CardResult <> 0 Then
        MsgBox "Read light On fail"
        End
    End If
                           
    ClosePipe
                      
    If rv0 = 1 Then
        CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
        Call MsecDelay(0.02)
                        
        If LightOn <> 127 Then
            Call MsecDelay(0.1)
            CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
            Call MsecDelay(0.02)
                            
            If LightOn <> 127 Then
                Call MsecDelay(0.1)
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
            End If
        End If
                        
        rv0 = CBWTest_New(0, 1, ChipString)
    End If
                    
    If rv0 = 1 Then
        rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
        
        If rv0 <> 1 Then
            rv0 = 2
            Tester.Print "SD bus width Fail"
        End If
    End If
                      
    If rv0 <> 1 Then ' for E55 command
        rv0 = Read_SD_SpeedE55(0, 0, 64, "8Bits")
        If rv0 <> 1 Then
            rv0 = 2
            Tester.Print "SD bus width Fail"
        End If
                        
        If rv0 = 1 Then
            CPRMMODE = 1
        End If
    End If
                    
    ClosePipe
                    
    Tester.Print "rv0="; rv0
                     
    If (rv0 = 0) Or (rv0 = 2) Then 'for OSE request
        rv0 = 2
        GoTo AU6371ELResult
    End If
                     
'=======================================================================================
    'SD 128Sector R / W
'=======================================================================================
                      
    TmpLBA = LBA
    rv1 = 0
    LBA = LBA + 199
                            
    ClosePipe
                        
    rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             
    If rv1 <> 1 Then
        LBA = TmpLBA
        GoTo AU6371ELResult
    End If
                      
'=======================================================================================
                     
    Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
    Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                
'===============================================
'  CF Card test
'================================================
               
    'rv1 = rv0  '----------- AU6371S3 dp not have CF slot
    'Call LabelMenu(1, rv1, rv0)
            
    'Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                
'===============================================
'  SMC Card test  : stop these test for card not enough
'================================================
              
    'AU6433 has no SMC slot
               
    rv2 = rv1   ' to complete the SMC asbolish
               
              
'===============================================
'  XD Card test
'================================================
                
    CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
    If CardResult <> 0 Then
        MsgBox "Set XD Card Detect On Fail"
        End
    End If
                  
    Call MsecDelay(0.1)
                
    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                     
    rv4 = rv3  'AU6344 has no MS slot pin
               
    'Call LabelMenu(2, rv4, rv3)
                 
    'Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
               
                
                
'===============================================
'  MS Pro Card test
'================================================
                
    CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
    Call MsecDelay(0.1)
               
    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
    Call MsecDelay(0.1)
                
    OpenPipe
    rv5 = ReInitial(0)
    ClosePipe
                
    rv5 = CBWTest_New(0, rv4, ChipString)
                
                
    If CPRMMODE = 0 Then  ' for E54 Before
        If rv5 = 1 Then
            rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                        
            If rv5 <> 1 Then
                rv5 = 2
                Tester.Print "MS bus width Fail"
            End If
        End If
    Else             ' for AU6433E55 after
                
        If rv5 = 1 Then
            rv5 = Read_MS_Speed_AU6476E55(0, 0, 64, "4Bits")
                If rv5 <> 1 Then
                    rv5 = 2
                    Tester.Print "MS bus width Fail"
                End If
        End If
    End If
                
    Call LabelMenu(31, rv5, rv4)
                
    Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    ClosePipe
                 
    CardResult = DO_WritePort(card, Channel_P1A, &H7F)   'Check NB mode
    Call MsecDelay(0.2)
                
    If GetDeviceName(ChipString) <> "" Then
        rv0 = 0
        Tester.Print "NB Mode test fail!"
        GoTo AU6371ELResult
    End If
                
    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                    
    Tester.Print "rv0="; rv0
                     
    If rv0 <> 0 Then
        If LightOn <> &H7F Or LightOff <> &HFF Then
            Tester.Print "LightON="; LightOn
            Tester.Print "LightOFF="; LightOff
            UsbSpeedTestResult = GPO_FAIL
            rv0 = 3
        End If
    End If
                
    Call LabelMenu(31, rv5, rv0)
                  
                                    
AU6371ELResult:

    CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
    Call PowerSet2(1, "0.0", "0.2", 1, "0.0", "0.2", 1)
    
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
Public Sub AU6433DLF00TestSub()

'2011/7/15 for GSMC & CSMC FT3

Dim TmpLBA As Long
Dim i As Integer
Dim HV_Flag As Boolean
Dim HV_Result As String
Dim LV_Result As String


If PCI7248InitFinish = 0 Then
    PCI7248ExistAU6254
    Call SetTimer_1ms
    Call MsecDelay(0.2)
    CardResult = DO_WritePort(card, Channel_P1C, &HFF)
End If


'Call PowerSet2(1, "3.3", "0.2", 1, "3.3", "0.2", 1)


'CardResult = DO_WritePort(card, Channel_P1C, &HFF)
                       

Tester.Print "Begin AU6433DL HV+LV Test..."

'==================================================================
'
'  this code come from AU6433DLF22TestSub
'  Purpose to solve GPON7 Detect issue
'
'===================================================================


    Dim ChipString As String
                
    Dim AU6371EL_SD As Byte
    Dim AU6371EL_CF As Byte
    Dim AU6371EL_XD As Byte
    Dim AU6371EL_MS As Byte
    Dim AU6371EL_MSP  As Byte
    Dim AU6371EL_BootTime As Single
    Dim CPRMMODE As Byte
    OldChipName = ""
               
Routine_Label:
                 
    ' initial condition
                
    AU6371EL_SD = 1
    AU6371EL_CF = 2
    AU6371EL_XD = 8
    AU6371EL_MS = 32
    AU6371EL_MSP = 64
                    
    AU6371EL_BootTime = 0.6
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
                  
    ChipString = "vid_058f"
'================================================
                   
Routine_Label_6435DLF00:
'=======================================================================================
    'SD R / W
'=======================================================================================
' set SD card detect down


    If (HV_Flag = False) Then
        Call PowerSet2(1, "3.6", "0.2", 1, "3.6", "0.2", 1)
        Tester.Print "Begin HV Test ..."
    Else
        Call PowerSet2(1, "3.05", "0.2", 1, "3.05", "0.2", 1)
        Tester.Print vbCrLf & "Begin LV Test ..."
        Call MsecDelay(0.2)
    End If
    
    CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
    If CardResult <> 0 Then
        MsgBox "Set SD Card Detect Down Fail"
        End
    End If
                     
    Call MsecDelay(0.1)
    rv0 = WaitDevOn(ChipString)
    Call MsecDelay(0.1)
                       
    If CardResult <> 0 Then
        MsgBox "Read light On fail"
        End
    End If
                           
    ClosePipe
                      
    If rv0 = 1 Then
        CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
        Call MsecDelay(0.02)
                        
        If LightOn <> 127 Then
            Call MsecDelay(0.1)
            CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
            Call MsecDelay(0.02)
                            
            If LightOn <> 127 Then
                Call MsecDelay(0.1)
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
            End If
        End If
                        
        rv0 = CBWTest_New(0, 1, ChipString)
    End If
                    
    If rv0 = 1 Then
        rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
        
        If rv0 <> 1 Then
            rv0 = 2
            Tester.Print "SD bus width Fail"
        End If
    End If
                      
    If rv0 <> 1 Then ' for E55 command
        rv0 = Read_SD_SpeedE55(0, 0, 64, "8Bits")
        If rv0 <> 1 Then
            rv0 = 2
            Tester.Print "SD bus width Fail"
        End If
                        
        If rv0 = 1 Then
            CPRMMODE = 1
        End If
    End If
                    
    ClosePipe
                    
                     
'=======================================================================================
    'SD 128Sector R / W
'=======================================================================================
                      
    TmpLBA = LBA
    LBA = LBA + 199
                            
    ClosePipe
                        
    If rv0 = 1 Then
        rv0 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
    End If
                  
'=======================================================================================
                     
    Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
    Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                
'===============================================
'  CF Card test
'================================================
               
    rv1 = rv0  '----------- AU6371S3 dp not have CF slot
    'Call LabelMenu(1, rv1, rv0)
            
    'Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                
'===============================================
'  SMC Card test  : stop these test for card not enough
'================================================
              
    'AU6433 has no SMC slot
               
    rv2 = rv1   ' to complete the SMC asbolish
               
              
'===============================================
'  XD Card test
'================================================
                
    CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
    If CardResult <> 0 Then
        MsgBox "Set XD Card Detect On Fail"
        End
    End If
                  
    Call MsecDelay(0.1)
                
    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                     
    rv4 = rv3  'AU6344 has no MS slot pin
               
    'Call LabelMenu(2, rv4, rv3)
                 
    'Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
               
                
                
'===============================================
'  MS Pro Card test
'================================================
                
    CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
    Call MsecDelay(0.1)
               
    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
    Call MsecDelay(0.1)
                
    OpenPipe
    rv5 = ReInitial(0)
    ClosePipe
                
    rv5 = CBWTest_New(0, rv4, ChipString)
                
                
    If CPRMMODE = 0 Then  ' for E54 Before
        If rv5 = 1 Then
            rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                        
            If rv5 <> 1 Then
                rv5 = 2
                Tester.Print "MS bus width Fail"
            End If
        End If
    Else             ' for AU6433E55 after
                
        If rv5 = 1 Then
            rv5 = Read_MS_Speed_AU6476E55(0, 0, 64, "4Bits")
                If rv5 <> 1 Then
                    rv5 = 2
                    Tester.Print "MS bus width Fail"
                End If
        End If
    End If
                
    Call LabelMenu(31, rv5, rv4)
                
    Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    ClosePipe
                 
    CardResult = DO_WritePort(card, Channel_P1A, &H7F)   'Check NB mode
    Call MsecDelay(0.2)
                
    If GetDeviceName(ChipString) <> "" Then
        rv0 = 0
        Tester.Print "NB Mode test fail!"
        GoTo AU6371ELResult
    End If
                
    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                    
    Tester.Print "rv0="; rv0
                     
    If rv0 <> 0 Then
        If LightOn <> &H7F Or LightOff <> &HFF Then
            Tester.Print "LightON="; LightOn
            Tester.Print "LightOFF="; LightOff
            UsbSpeedTestResult = GPO_FAIL
            rv0 = 3
        End If
    End If
                
    Call LabelMenu(31, rv5, rv0)
                  
                                    
AU6371ELResult:

    CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
    Call PowerSet2(1, "0.0", "0.2", 1, "0.0", "0.2", 1)
    
    If HV_Flag = False Then
                    If rv0 * rv1 * rv2 * rv3 * rv4 * rv5 = 0 Then
                        HV_Result = "Bin2"
                        Tester.Print "HV Unknow"
                    ElseIf rv0 * rv1 * rv2 * rv3 * rv4 * rv5 <> 1 Then
                        HV_Result = "Fail"
                        Tester.Print "HV Fail"
                    ElseIf rv0 * rv1 * rv2 * rv3 * rv4 * rv5 = 1 Then
                        HV_Result = "PASS"
                        Tester.Print "HV PASS"
                    End If
                    ReaderExist = 0
                    HV_Flag = True
                    Call MsecDelay(0.2)
                    GoTo Routine_Label_6435DLF00
                Else
                    If rv0 * rv1 * rv2 * rv3 * rv4 * rv5 = 0 Then
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
    
    
End Sub
Public Sub AU6433DLF3CTestSub()

'2011/7/13 for CSMC (Just O/S pattern different)

Dim TmpLBA As Long
Dim i As Integer

If PCI7248InitFinish = 0 Then
    PCI7248ExistAU6254
    Call SetTimer_1ms
End If


Call PowerSet2(1, "3.3", "0.2", 1, "3.3", "0.2", 1)

OS_Result = 0
rv0 = 0

CardResult = DO_WritePort(card, Channel_P1C, &H0)
                 
MsecDelay (0.3)

OpenShortTest_Result

If OS_Result <> 1 Then
    rv0 = 0                 'OS Fail
    GoTo AU6371ELResult
End If

CardResult = DO_WritePort(card, Channel_P1C, &HFF)
                       

Tester.Print "AU6433DL Using Internal LDO Output..."

'==================================================================
'
'  this code come from AU6433DLF22TestSub
'  Purpose to solve GPON7 Detect issue
'
'===================================================================


    Dim ChipString As String
                
    Dim AU6371EL_SD As Byte
    Dim AU6371EL_CF As Byte
    Dim AU6371EL_XD As Byte
    Dim AU6371EL_MS As Byte
    Dim AU6371EL_MSP  As Byte
    Dim AU6371EL_BootTime As Single
    Dim CPRMMODE As Byte
    OldChipName = ""
               
                 
                 
    ' initial condition
                
    AU6371EL_SD = 1
    AU6371EL_CF = 2
    AU6371EL_XD = 8
    AU6371EL_MS = 32
    AU6371EL_MSP = 64
                    
    AU6371EL_BootTime = 0.6
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
                  
    ChipString = "vid"
'================================================
                   

'=======================================================================================
    'SD R / W
'=======================================================================================
' set SD card detect down
    CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
    If CardResult <> 0 Then
        MsgBox "Set SD Card Detect Down Fail"
        End
    End If
                     
    Call MsecDelay(0.2)
    rv0 = WaitDevOn(ChipString)
    Call MsecDelay(0.1)
                       
    If CardResult <> 0 Then
        MsgBox "Read light On fail"
        End
    End If
                           
    ClosePipe
                      
    If rv0 = 1 Then
        CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
        Call MsecDelay(0.02)
                        
        If LightOn <> 127 Then
            Call MsecDelay(0.1)
            CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
            Call MsecDelay(0.02)
                            
            If LightOn <> 127 Then
                Call MsecDelay(0.1)
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
            End If
        End If
                        
        rv0 = CBWTest_New(0, 1, ChipString)
    End If
                    
    If rv0 = 1 Then
        rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
        
        If rv0 <> 1 Then
            rv0 = 2
            Tester.Print "SD bus width Fail"
        End If
    End If
                      
    If rv0 <> 1 Then ' for E55 command
        rv0 = Read_SD_SpeedE55(0, 0, 64, "8Bits")
        If rv0 <> 1 Then
            rv0 = 2
            Tester.Print "SD bus width Fail"
        End If
                        
        If rv0 = 1 Then
            CPRMMODE = 1
        End If
    End If
                    
    ClosePipe
                    
    Tester.Print "rv0="; rv0
                     
    If (rv0 = 0) Or (rv0 = 2) Then 'for OSE request
        rv0 = 2
        GoTo AU6371ELResult
    End If
                     
'=======================================================================================
    'SD 128Sector R / W
'=======================================================================================
                      
    TmpLBA = LBA
    rv1 = 0
    LBA = LBA + 199
                            
    ClosePipe
                        
    rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             
    If rv1 <> 1 Then
        LBA = TmpLBA
        GoTo AU6371ELResult
    End If
                      
'=======================================================================================
                     
    Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
    Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                
'===============================================
'  CF Card test
'================================================
               
    'rv1 = rv0  '----------- AU6371S3 dp not have CF slot
    'Call LabelMenu(1, rv1, rv0)
            
    'Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                
'===============================================
'  SMC Card test  : stop these test for card not enough
'================================================
              
    'AU6433 has no SMC slot
               
    rv2 = rv1   ' to complete the SMC asbolish
               
              
'===============================================
'  XD Card test
'================================================
                
    CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
    If CardResult <> 0 Then
        MsgBox "Set XD Card Detect On Fail"
        End
    End If
                  
    Call MsecDelay(0.1)
                
    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                     
    rv4 = rv3  'AU6344 has no MS slot pin
               
    'Call LabelMenu(2, rv4, rv3)
                 
    'Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
               
                
                
'===============================================
'  MS Pro Card test
'================================================
                
    CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
    Call MsecDelay(0.1)
               
    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
    Call MsecDelay(0.1)
                
    OpenPipe
    rv5 = ReInitial(0)
    ClosePipe
                
    rv5 = CBWTest_New(0, rv4, ChipString)
                
                
    If CPRMMODE = 0 Then  ' for E54 Before
        If rv5 = 1 Then
            rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                        
            If rv5 <> 1 Then
                rv5 = 2
                Tester.Print "MS bus width Fail"
            End If
        End If
    Else             ' for AU6433E55 after
                
        If rv5 = 1 Then
            rv5 = Read_MS_Speed_AU6476E55(0, 0, 64, "4Bits")
                If rv5 <> 1 Then
                    rv5 = 2
                    Tester.Print "MS bus width Fail"
                End If
        End If
    End If
                
    Call LabelMenu(31, rv5, rv4)
                
    Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    ClosePipe
                 
    CardResult = DO_WritePort(card, Channel_P1A, &H7F)   'Check NB mode
    Call MsecDelay(0.2)
                
    If GetDeviceName(ChipString) <> "" Then
        rv0 = 0
        Tester.Print "NB Mode test fail!"
        GoTo AU6371ELResult
    End If
                
    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                    
    Tester.Print "rv0="; rv0
                     
    If rv0 <> 0 Then
        If LightOn <> &H7F Or LightOff <> &HFF Then
            Tester.Print "LightON="; LightOn
            Tester.Print "LightOFF="; LightOff
            UsbSpeedTestResult = GPO_FAIL
            rv0 = 3
        End If
    End If
                
    Call LabelMenu(31, rv5, rv0)
                  
                                    
AU6371ELResult:

    CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
    Call PowerSet2(1, "0.0", "0.2", 1, "0.0", "0.2", 1)
    
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
Public Sub AU6433EFF35TestSub()

If PCI7248InitFinish = 0 Then
    PCI7248ExistAU6254
    Call SetTimer_1ms
End If

OS_Result = 0
rv0 = 0

CardResult = DO_WritePort(card, Channel_P1C, &H0)   'Set Switch connect to OS Board
                 
MsecDelay (0.3)

OpenShortTest_Result

If OS_Result <> 1 Then
    rv0 = 0                 'OS Fail
    GoTo AU6371ELResult
End If

TestResult = ""

CardResult = DO_WritePort(card, Channel_P1C, &HFF)   'Set Switch connect to FT Module
                 
'MsecDelay (0.3)

  Tester.Print "AU6433EF : NB mode test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim CPRMMODE As Byte
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
                ChipString = "vid"
                 If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                  Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1.3)
                     
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
                      
                         If rv0 <> 1 Then ' for E55 command
                          rv0 = Read_SD_SpeedE55(0, 0, 64, "8Bits")
                          If rv0 <> 1 Then
                          rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                          If rv0 = 1 Then
                            CPRMMODE = 1
                          End If
                      End If
                      
                      
                      ClosePipe
                      
                      Tester.Print "rv0="; rv0
                     
                         If rv0 <> 0 Then
                           If LightOn <> &HBF Or LightOff <> &HFF Then
                           Tester.Print "LightON="; LightOn
                         Tester.Print "LightOFF="; LightOff
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
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
                 Call MsecDelay(0.1)
               
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
                 Call MsecDelay(0.1)
                  OpenPipe
                  rv5 = ReInitial(0)
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                If CPRMMODE = 0 Then  ' for E54 Before
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                 Else             ' for AU6433E55 after
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476E55(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
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
Public Sub AU6433EFF36TestSub()

If PCI7248InitFinish = 0 Then
    PCI7248ExistAU6254
    Call SetTimer_1ms
End If

OS_Result = 0
rv0 = 0

CardResult = DO_WritePort(card, Channel_P1C, &H0)   'Set Switch connect to OS Board
                 
MsecDelay (0.3)

OpenShortTest_Result_AU6433EFF35

If OS_Result <> 1 Then
    rv0 = 0                 'OS Fail
    GoTo AU6371ELResult
End If

TestResult = ""

CardResult = DO_WritePort(card, Channel_P1C, &HFF)   'Set Switch connect to FT Module
                 
'MsecDelay (0.3)

  Tester.Print "AU6433EF : NB mode test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim CPRMMODE As Byte
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
                ChipString = "vid"
                 If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                  Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1.3)
                     
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
                      
                         If rv0 <> 1 Then ' for E55 command
                          rv0 = Read_SD_SpeedE55(0, 0, 64, "8Bits")
                          If rv0 <> 1 Then
                          rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                          If rv0 = 1 Then
                            CPRMMODE = 1
                          End If
                      End If
                      
                      
                      ClosePipe
                      
                      Tester.Print "rv0="; rv0
                     
                         If rv0 <> 0 Then
                           If LightOn <> &HBF Or LightOff <> &HFF Then
                           Tester.Print "LightON="; LightOn
                         Tester.Print "LightOFF="; LightOff
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
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
                 Call MsecDelay(0.1)
               
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
                 Call MsecDelay(0.1)
                  OpenPipe
                  rv5 = ReInitial(0)
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                If CPRMMODE = 0 Then  ' for E54 Before
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                 Else             ' for AU6433E55 after
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476E55(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
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

Public Sub AU6433EFF25TestSub()
      
  Tester.Print "AU6433EF : NB mode test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim CPRMMODE As Byte
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                    If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
                ChipString = "vid"
                 If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                  Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1.3)
                     
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
                      
                         If rv0 <> 1 Then ' for E55 command
                          rv0 = Read_SD_SpeedE55(0, 0, 64, "8Bits")
                          If rv0 <> 1 Then
                          rv0 = 2
                          Tester.Print "SD bus width Fail"
                          End If
                          If rv0 = 1 Then
                            CPRMMODE = 1
                          End If
                      End If
                      
                      
                      ClosePipe
                      
                      Tester.Print "rv0="; rv0
                     
                         If rv0 <> 0 Then
                           If LightOn <> &HBF Or LightOff <> &HFF Then
                           Tester.Print "LightON="; LightOn
                         Tester.Print "LightOFF="; LightOff
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
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
                 Call MsecDelay(0.1)
               
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
                 Call MsecDelay(0.1)
                  OpenPipe
                  rv5 = ReInitial(0)
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                If CPRMMODE = 0 Then  ' for E54 Before
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                 Else             ' for AU6433E55 after
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476E55(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
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

Public Sub AU6433EFF26TestSub()



  Tester.Print "AU6433EF : NB mode test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                Dim CPRMMODE As Byte
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
               
                ChipString = "vid_058f"
                 
                '===============================================
                '  SD Card test
                '
                  
                 '===========================================
                 'NO card test
                 '============================================
                
                ' set SD card detect down
                CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                      
                If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect Down Fail"
                    End
                End If
                     
                Call MsecDelay(0.3)
                          
                CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                End If
                
                
                ' set SD card detect down
                CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect Down Fail"
                    End
                End If
                
                Call MsecDelay(0.2)
                rv0 = WaitDevOn(ChipString)
                Call MsecDelay(0.1)
                     
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                If CardResult <> 0 Then
                    MsgBox "Read light On fail"
                    End
                End If
                           
                ClosePipe
                      
                rv0 = CBWTest_New(0, 1, ChipString)
                If rv0 <> 1 Then
                    ClosePipe
                    GoTo AU6371ELResult
                End If
                
                If rv0 = 1 Then
                    rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
                    End If
                End If
                      
                If rv0 <> 1 Then ' for E55 command
                    rv0 = Read_SD_SpeedE55(0, 0, 64, "8Bits")
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
                    End If
                    
                    If rv0 = 1 Then
                        CPRMMODE = 1
                    End If
                End If
                      
                      
                ClosePipe
                      
                Tester.Print "rv0="; rv0
                     
                If rv0 <> 0 Then
                    If LightOn <> &HBF Or LightOff <> &HFF Then
                        Tester.Print "LightON="; LightOn
                        Tester.Print "LightOFF="; LightOff
                        UsbSpeedTestResult = GPO_FAIL
                        rv0 = 3
                    End If
                End If
                    
                     
                Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  CF Card test
                '================================================
               
                rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                End If
                  
                Call MsecDelay(0.1)
                CardResult = DO_WritePort(card, Channel_P1A, &H77)
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
                   
                     
                rv4 = rv3  'AU6344 has no MS slot pin
               
                'Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                CardResult = DO_WritePort(card, Channel_P1A, &H57)
                Call MsecDelay(0.1)
                CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                 
                Call MsecDelay(0.1)
                OpenPipe
                rv5 = ReInitial(0)
                ClosePipe
                
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                If CPRMMODE = 0 Then  ' for E54 Before
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                 Else             ' for AU6433E55 after
                        If rv5 = 1 Then
                           rv5 = Read_MS_Speed_AU6476E55(0, 0, 64, "4Bits")
                           If rv5 <> 1 Then
                              rv5 = 2
                              Tester.Print "MS bus width Fail"
                           End If
                        End If
                End If
                   
                
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                 
                If rv5 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                    Call MsecDelay(0.2)
                    If GetDeviceName(ChipString) <> "" Then
                        rv0 = 0
                        Tester.Print "NB-MODE Fail"
                    Else
                        Tester.Print "NB-MODE PASS"
                    End If
                End If
                 
                  
AU6371ELResult:
                CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
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
Public Sub AU6433EFF23TestSub()
      
  Tester.Print "AU6433EF : NB mode test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
                ChipString = "vid"
                 If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                  Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1#)
                     
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
                      
                      Tester.Print "rv0="; rv0
                     
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                          Tester.Print "LightON="; LightOn
                          Tester.Print "LightOFF="; LightOff
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
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
                 Call MsecDelay(0.1)
               
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
                 Call MsecDelay(0.1)
                  OpenPipe
                  rv5 = ReInitial(0)
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                If rv5 = 1 Then
                   rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                   If rv5 <> 1 Then
                      rv5 = 2
                      Tester.Print "MS bus width Fail"
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
Public Sub AU6433EFF22TestSub()
      
  Tester.Print "AU6433EF : NB mode test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
                ChipString = "vid"
                 If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                  Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1#)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
                      
                      Tester.Print "rv0="; rv0
                     
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                          Tester.Print "LightON="; LightOn
                          Tester.Print "LightOFF="; LightOff
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
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
                 Call MsecDelay(0.1)
               
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
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
Public Sub AU6433LFF33TestSub()
      
             

If PCI7248InitFinish = 0 Then
    PCI7248ExistAU6254
    Call SetTimer_1ms
End If

OS_Result = 0
rv0 = 0

CardResult = DO_WritePort(card, Channel_P1C, &H0)   'Set Switch connect to OS Board
                 
MsecDelay (0.3)

OpenShortTest_Result

If OS_Result <> 1 Then
    rv0 = 0                 'OS Fail
    GoTo AU6371ELResult
End If

TestResult = ""

CardResult = DO_WritePort(card, Channel_P1C, &HFF)   'Set Switch connect to FT Module
                 
'MsecDelay (0.3)
      
      
      
  Tester.Print "AU6433LF : NB mode test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
                ChipString = "vid"
                 If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 '================================================
                 '  CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
                      
                  
                 '  If CardResult <> 0 Then
                 '   MsgBox "Read light off fail"
                 '   End
                 '  End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                  Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1#)
                     
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
                      
                      Tester.Print "rv0="; rv0
                     
                       ' If rv0 <> 0 Then
                       '   If LightON <> &HBF Or LightOFF <> &HFF Then
                       '   Tester.Print "LightON="; LightON
                       '   Tester.Print "LightOFF="; LightOFF
                       '   UsbSpeedTestResult = GPO_FAIL
                       '   rv0 = 3
                       '   End If
                       ' End If
                    
                     
                     
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
                 Call MsecDelay(0.1)
               
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
                 Call MsecDelay(0.1)
                  OpenPipe
                  rv5 = ReInitial(0)
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                      If rv5 = 1 Then
                          rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                          If rv5 <> 1 Then
                          rv5 = 2
                          Tester.Print "MS bus width Fail"
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

Public Sub AU6433LFF23TestSub()
      
  Tester.Print "AU6433LF : NB mode test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
                ChipString = "vid"
                 If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 '================================================
                 '  CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
                      
                  
                 '  If CardResult <> 0 Then
                 '   MsgBox "Read light off fail"
                 '   End
                 '  End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                  Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1#)
                     
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
                      
                      Tester.Print "rv0="; rv0
                     
                       ' If rv0 <> 0 Then
                       '   If LightON <> &HBF Or LightOFF <> &HFF Then
                       '   Tester.Print "LightON="; LightON
                       '   Tester.Print "LightOFF="; LightOFF
                       '   UsbSpeedTestResult = GPO_FAIL
                       '   rv0 = 3
                       '   End If
                       ' End If
                    
                     
                     
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
                 Call MsecDelay(0.1)
               
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
                 Call MsecDelay(0.1)
                  OpenPipe
                  rv5 = ReInitial(0)
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                      If rv5 = 1 Then
                          rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                          If rv5 <> 1 Then
                          rv5 = 2
                          Tester.Print "MS bus width Fail"
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

Public Sub AU6433LFF22TestSub()
      
  Tester.Print "AU6433LF : NB mode test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
                ChipString = "vid"
                 If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 '================================================
                 '  CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
                      
                  
                 '  If CardResult <> 0 Then
                 '   MsgBox "Read light off fail"
                 '   End
                 '  End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                  Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1#)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
                      
                      Tester.Print "rv0="; rv0
                     
                       ' If rv0 <> 0 Then
                       '   If LightON <> &HBF Or LightOFF <> &HFF Then
                       '   Tester.Print "LightON="; LightON
                       '   Tester.Print "LightOFF="; LightOFF
                       '   UsbSpeedTestResult = GPO_FAIL
                       '   rv0 = 3
                       '   End If
                       ' End If
                    
                     
                     
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
                 Call MsecDelay(0.1)
               
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
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


Public Sub AU6433IFF23TestSub()
      
  Tester.Print "AU6433IF : NB mode test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
                ChipString = "vid"
                 If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                  Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1#)
                     
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
                      
                      Tester.Print "rv0="; rv0
                     
                       ' If rv0 <> 0 Then
                       '   If LightON <> &HBF Or LightOFF <> &HFF Then
                       '   Tester.Print "LightON="; LightON
                       '   Tester.Print "LightOFF="; LightOFF
                       '   UsbSpeedTestResult = GPO_FAIL
                       '   rv0 = 3
                       '   End If
                       ' End If
                    
                     
                     
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
                 Call MsecDelay(0.1)
               
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
                 Call MsecDelay(0.1)
                  OpenPipe
                  rv5 = ReInitial(0)
                
                  
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                   If rv5 = 1 Then
                         rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                         If rv5 <> 1 Then
                         rv5 = 2
                         Tester.Print "MS bus width Fail"
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
Public Sub AU6433IFF22TestSub()
      
  Tester.Print "AU6433IF : NB mode test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
                ChipString = "vid"
                 If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                  Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1#)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
                      
                      Tester.Print "rv0="; rv0
                     
                       ' If rv0 <> 0 Then
                       '   If LightON <> &HBF Or LightOFF <> &HFF Then
                       '   Tester.Print "LightON="; LightON
                       '   Tester.Print "LightOFF="; LightOFF
                       '   UsbSpeedTestResult = GPO_FAIL
                       '   rv0 = 3
                       '   End If
                       ' End If
                    
                     
                     
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                  '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                 End If
                  
                  
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
              
                 Call MsecDelay(0.1)
               
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
               
                 
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
Public Sub AU6433EFTest()
      
  Tester.Print "AU6433EF : NB mode test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
                ChipString = "vid"
                 If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                  Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1#)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
                      
                      Tester.Print "rv0="; rv0
                     
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                          Tester.Print "LightON="; LightOn
                          Tester.Print "LightOFF="; LightOff
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
               
                  rv1 = 1  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                 Call LabelMenu(1, rv1, rv0)
            
                      Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = 1   ' to complete the SMC asbolish
               
              
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
                 
                  
                ReaderExist = 0
                 
                ClosePipe
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  MS Card test
                '================================================
                   
                     
                     rv4 = 1  'AU6344 has no MS slot pin
               
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
                 
                ReaderExist = 0
                
                
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
Public Sub AU6433GSF23TestSub()
      
  Tester.Print "AU6433GS : NB mode test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'  this code from AU6433EFTestSub at AU6433EFF21
'  ' add : NB mode test
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
                ChipString = "vid"
                 If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                  Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1#)
                     
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
                      
                      Tester.Print "rv0="; rv0
                      
                       ' this chip , do not need to test light
                       ' If rv0 <> 0 Then
                       '   If LightON <> &HBF Or LightOFF <> &HFF Then
                       '   Tester.Print "LightON="; LightON
                       '   Tester.Print "LightOFF="; LightOFF
                       '   UsbSpeedTestResult = GPO_FAIL
                       '   rv0 = 3
                       '   End If
                       ' End If
                    
                     
                     
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                 Call LabelMenu(1, rv1, rv0)
            
                '      Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76)
                  
                   
               
                  
                 Call MsecDelay(0.01)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                 
                Call MsecDelay(0.01)
                 
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
              '       Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
                
              
                
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                Call MsecDelay(0.1)
                
                  OpenPipe
                 rv5 = ReInitial(0)
                  
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                If rv5 = 1 Then
                   rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                   If rv5 <> 1 Then
                     rv5 = 2
                     Tester.Print "MS bus width Fail"
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
Public Sub AU6433JSF28TestSub()

  
             
Call PowerSet2(1, "3.3", "0.09", 1, "2.1", "0.09", 1)
  Tester.Print "AU6433JS : NB mode test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'  this code from AU6433EFTestSub at AU6433EFF21
'  ' add : NB mode test
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
               If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
               
             '  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
             '   CardResult = DO_WritePort(card, Channel_P1B, &HFF)
               
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
                ChipString = "vid"
                 If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                  Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1#)
                     
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
                      
                      Tester.Print "rv0="; rv0
                      
                       ' this chip , do not need to test light
                       ' If rv0 <> 0 Then
                       '   If LightON <> &HBF Or LightOFF <> &HFF Then
                       '   Tester.Print "LightON="; LightON
                       '   Tester.Print "LightOFF="; LightOFF
                       '   UsbSpeedTestResult = GPO_FAIL
                       '   rv0 = 3
                       '   End If
                       ' End If
                    
                     
                     
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                 Call LabelMenu(1, rv1, rv0)
            
                '      Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76)
                  
                   
               
                  
                 Call MsecDelay(0.01)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                 
                Call MsecDelay(0.01)
                 
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
              '       Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
                
              
                
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                Call MsecDelay(0.1)
                
                  OpenPipe
                 rv5 = ReInitial(0)
                  
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                If rv5 = 1 Then
                   rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                   If rv5 <> 1 Then
                     rv5 = 2
                     Tester.Print "MS bus width Fail"
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
Public Sub AU6433JSF29TestSub()


Dim TmpLBA As Long
Dim i As Integer
             
Call PowerSet2(1, "3.3", "0.09", 1, "2.1", "0.09", 1)
  Tester.Print "AU6433JS : NB mode test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'  this code from AU6433EFTestSub at AU6433EFF21
'  ' add : NB mode test
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
               If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
               
             '  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
             '   CardResult = DO_WritePort(card, Channel_P1B, &HFF)
               
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
                ChipString = "vid"
                 If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                  Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1#)
                     
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
                      
                      Tester.Print "rv0="; rv0
                      
                       ' this chip , do not need to test light
                       ' If rv0 <> 0 Then
                       '   If LightON <> &HBF Or LightOFF <> &HFF Then
                       '   Tester.Print "LightON="; LightON
                       '   Tester.Print "LightOFF="; LightOFF
                       '   UsbSpeedTestResult = GPO_FAIL
                       '   rv0 = 3
                       '   End If
                       ' End If
                    
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      
                     TmpLBA = LBA
                     'LBA = 99
                         'For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6371ELResult
                             End If
                         'Next
                      LBA = TmpLBA
                      
'=======================================================================================
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                 Call LabelMenu(1, rv1, rv0)
            
                '      Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76)
                  
                   
               
                  
                 Call MsecDelay(0.01)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                 
                Call MsecDelay(0.01)
                 
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
              '       Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
                
              
                
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                Call MsecDelay(0.1)
                
                  OpenPipe
                 rv5 = ReInitial(0)
                  
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                If rv5 = 1 Then
                   rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                   If rv5 <> 1 Then
                     rv5 = 2
                     Tester.Print "MS bus width Fail"
                   End If
                   
                End If
                   
                
                
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                 
                  CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                  
                  Call PowerSet2(1, "0.0", "0.09", 1, "0.0", "0.09", 1)

                
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
Public Sub AU6433JSF2ATestSub()


Dim TmpLBA As Long
Dim i As Integer
             
'==================================================================
'
'  this code come from AU6371ELTestSub
'  this code from AU6433EFTestSub at AU6433EFF21
'   add : NB mode test
'
'===================================================================


Dim ChipString As String
Dim AU6371EL_SD As Byte
Dim AU6371EL_CF As Byte
Dim AU6371EL_XD As Byte
Dim AU6371EL_MS As Byte
Dim AU6371EL_MSP  As Byte
Dim AU6371EL_BootTime As Single
OldChipName = ""
                 
' initial condition
                
    AU6371EL_SD = 1
    AU6371EL_CF = 2
    AU6371EL_XD = 8
    AU6371EL_MS = 32
    AU6371EL_MSP = 64
    AU6371EL_BootTime = 0.6
                     
                    
    If PCI7248InitFinish = 0 Then
        PCI7248Exist
    End If
               
    'result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
    'CardResult = DO_WritePort(card, Channel_P1B, &HFF)
               
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
        
    Call MsecDelay(0.1)
                  
    ChipString = "vid_058f"
               
    '===============================================
    '  SD Card test
    '===============================================
                
    ' set SD card detect down
    
    CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
    If CardResult <> 0 Then
        MsgBox "Set SD Card Detect Down Fail"
        End
    End If
    
    Call MsecDelay(0.2)
    rv0 = WaitDevOn(ChipString)
    Call MsecDelay(0.1)
                     
                           
    ClosePipe
    
    If rv0 = 1 Then
        rv0 = CBWTest_New(0, rv0, ChipString)
        
        If rv0 = 1 Then
            rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
            If rv0 <> 1 Then
                rv0 = 2
                Tester.Print "SD bus width Fail"
            End If
        End If
        
        ClosePipe
    End If
                     
'=======================================================================================
    'SD R/W 64K
'=======================================================================================
                      
    If rv0 = 1 Then
        TmpLBA = LBA
        LBA = LBA + 199
    
        ClosePipe
    
        rv0 = CBWTest_New_128_Sector_AU6377(0, rv0)  ' write
        If rv0 <> 1 Then
            LBA = TmpLBA
            Tester.Print "R/W 64K Fail ..."
        End If
        LBA = TmpLBA
        ClosePipe
    End If
    
    Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
    Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    '===============================================
    '  CF Card test
    '================================================
               
    rv1 = rv0  '----------- AU6433JS no CF slot
    'Call LabelMenu(1, rv1, rv0)
    'Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    
    '===============================================
    '  SMC Card test  : stop these test for card not enough
    '================================================
              
    'AU6433JS no SMC slot
               
    rv2 = rv1   ' to complete the SMC asbolish
    'Call LabelMenu(1, rv2, rv1)
    'Tester.Print rv2, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
               
              
    '===============================================
    '  XD Card test
    '================================================
    CardResult = DO_WritePort(card, Channel_P1A, &H76)
    Call MsecDelay(0.01)
    CardResult = DO_WritePort(card, Channel_P1A, &H77)
    Call MsecDelay(0.01)
                 
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
                     
    rv4 = rv3  'AU6433JS no MS slot
               
    'Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    '===============================================
    '  MS Pro Card test
    '================================================
                
    CardResult = DO_WritePort(card, Channel_P1A, &H57)
    Call MsecDelay(0.1)
    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
    Call MsecDelay(0.1)
    
    OpenPipe
    rv5 = ReInitial(0)
    ClosePipe
                
    rv5 = CBWTest_New(0, rv4, ChipString)
                
    If rv5 = 1 Then
        rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
        If rv5 <> 1 Then
            rv5 = 2
            Tester.Print "MS bus width Fail"
        End If
    End If
                
    Call LabelMenu(31, rv5, rv4)
    Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    ClosePipe
                 
    If rv5 = 1 Then
        CardResult = DO_WritePort(card, Channel_P1A, &H7F)
        Call MsecDelay(0.2)
        If GetDeviceName(ChipString) <> "" Then
            rv0 = 0
            Tester.Print "NBMD Test Fail ..."
        End If
    End If
    
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
Public Sub AU6433JSF3ATestSub()


Dim TmpLBA As Long
Dim i As Integer
             
'==================================================================
'
'  this code come from AU6371ELTestSub
'  this code from AU6433EFTestSub at AU6433EFF21
'   add : NB mode test
'
'===================================================================

Dim ChipString As String
Dim AU6371EL_SD As Byte
Dim AU6371EL_CF As Byte
Dim AU6371EL_XD As Byte
Dim AU6371EL_MS As Byte
Dim AU6371EL_MSP  As Byte
Dim AU6371EL_BootTime As Single
OldChipName = ""
                 
' initial condition
                
    AU6371EL_SD = 1
    AU6371EL_CF = 2
    AU6371EL_XD = 8
    AU6371EL_MS = 32
    AU6371EL_MSP = 64
    AU6371EL_BootTime = 0.6
                     
                    
    If PCI7248InitFinish = 0 Then
        PCI7248ExistAU6254
        Call SetTimer_1ms
    End If
    
    OS_Result = 0
    
    CardResult = DO_WritePort(card, Channel_P1C, &H0)   'Set Switch connect to OS Board
                 
    MsecDelay (0.2)
    OpenShortTest_Result

    If OS_Result <> 1 Then
        rv0 = 0                 'OS Fail
        GoTo AU6371ELResult
    End If

    TestResult = ""

    CardResult = DO_WritePort(card, Channel_P1C, &HFF)   'Set Switch connect to FT Module
    MsecDelay (0.2)
    
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
                 
    'CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
    ChipString = "vid_058f"
    
    '===============================================
    '  SD Card test
    '===============================================
                
    ' set SD card detect down
    
    CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
    If CardResult <> 0 Then
        MsgBox "Set SD Card Detect Down Fail"
        End
    End If
    
    Call MsecDelay(0.2)
    rv0 = WaitDevOn(ChipString)
    Call MsecDelay(0.1)
                     
                           
    ClosePipe
    
    If rv0 = 1 Then
        rv0 = CBWTest_New(0, rv0, ChipString)
        
        If rv0 = 1 Then
            rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
            If rv0 <> 1 Then
                rv0 = 2
                Tester.Print "SD bus width Fail"
            End If
        End If
        
        ClosePipe
    End If
                     
'=======================================================================================
    'SD R/W 64K
'=======================================================================================
                      
    If rv0 = 1 Then
        TmpLBA = LBA
        LBA = LBA + 199
    
        ClosePipe
        rv0 = CBWTest_New_128_Sector_AU6377(0, rv0)  ' write
        If rv0 <> 1 Then
            LBA = TmpLBA
            Tester.Print "R/W 64K Fail ..."
        End If
        
        LBA = TmpLBA
        ClosePipe
    End If
    
    Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
    Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    '===============================================
    '  CF Card test
    '================================================
               
    rv1 = rv0  '----------- AU6433JS no CF slot
    'Call LabelMenu(1, rv1, rv0)
    'Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    
    '===============================================
    '  SMC Card test  : stop these test for card not enough
    '================================================
              
    'AU6433JS no SMC slot
               
    rv2 = rv1   ' to complete the SMC asbolish
    'Call LabelMenu(1, rv2, rv1)
    'Tester.Print rv2, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
               
              
    '===============================================
    '  XD Card test
    '================================================
    CardResult = DO_WritePort(card, Channel_P1A, &H76)
    Call MsecDelay(0.01)
    CardResult = DO_WritePort(card, Channel_P1A, &H77)
    Call MsecDelay(0.01)
                 
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
                     
    rv4 = rv3  'AU6433JS no MS slot
               
    'Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    '===============================================
    '  MS Pro Card test
    '================================================
                
    CardResult = DO_WritePort(card, Channel_P1A, &H57)
    Call MsecDelay(0.02)
    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
    Call MsecDelay(0.02)
    
    OpenPipe
    rv5 = ReInitial(0)
    ClosePipe
                
    rv5 = CBWTest_New(0, rv4, ChipString)
                
    If rv5 = 1 Then
        rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
        If rv5 <> 1 Then
            rv5 = 2
            Tester.Print "MS bus width Fail"
        End If
    End If
                
    Call LabelMenu(31, rv5, rv4)
    Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    ClosePipe
                 
    If rv5 = 1 Then
        CardResult = DO_WritePort(card, Channel_P1A, &H7F)
        Call MsecDelay(0.2)
        If GetDeviceName(ChipString) <> "" Then
            rv0 = 0
            Tester.Print "NBMD Test Fail ..."
        End If
    End If
    
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
Public Sub AU6433JSF39TestSub()


Dim TmpLBA As Long
Dim i As Integer
             

               If PCI7248InitFinish = 0 Then
                  PCI7248ExistAU6254
                  Call SetTimer_1ms
               End If

Call PowerSet2(1, "3.3", "0.09", 1, "2.1", "0.09", 1)

OS_Result = 0
rv0 = 0

CardResult = DO_WritePort(card, Channel_P1C, &H0)   'Set Switch connect to OS Board
                 
MsecDelay (0.3)

OpenShortTest_Result

If OS_Result <> 1 Then
    rv0 = 0                 'OS Fail
    GoTo AU6371ELResult
End If

TestResult = ""

CardResult = DO_WritePort(card, Channel_P1C, &HFF)   'Set Switch connect to FT Module
                 
'MsecDelay (0.3)

  Tester.Print "AU6433JS : NB mode test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'  this code from AU6433EFTestSub at AU6433EFF21
'  ' add : NB mode test
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
               If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
               
             '  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
             '   CardResult = DO_WritePort(card, Channel_P1B, &HFF)
               
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
                ChipString = "vid"
                 If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                  Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1#)
                     
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
                      
                      Tester.Print "rv0="; rv0
                      
                       ' this chip , do not need to test light
                       ' If rv0 <> 0 Then
                       '   If LightON <> &HBF Or LightOFF <> &HFF Then
                       '   Tester.Print "LightON="; LightON
                       '   Tester.Print "LightOFF="; LightOFF
                       '   UsbSpeedTestResult = GPO_FAIL
                       '   rv0 = 3
                       '   End If
                       ' End If
                    
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      
                     TmpLBA = LBA
                     'LBA = 99
                         'For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6371ELResult
                             End If
                         'Next
                      LBA = TmpLBA
                      
'=======================================================================================
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                 Call LabelMenu(1, rv1, rv0)
            
                '      Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76)
                  
                   
               
                  
                 Call MsecDelay(0.01)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                 
                Call MsecDelay(0.01)
                 
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
              '       Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
                
              
                
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                Call MsecDelay(0.1)
                
                  OpenPipe
                 rv5 = ReInitial(0)
                  
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                If rv5 = 1 Then
                   rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                   If rv5 <> 1 Then
                     rv5 = 2
                     Tester.Print "MS bus width Fail"
                   End If
                   
                End If
                   
                
                
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                 
                  CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                  
                  Call PowerSet2(1, "0.0", "0.09", 1, "0.0", "0.09", 1)

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
Public Sub AU6433GSF22TestSub()
      
  Tester.Print "AU6433GS : NB mode test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'  this code from AU6433EFTestSub at AU6433EFF21
'  ' add : NB mode test
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
                ChipString = "vid"
                 If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                  Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1#)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
                      
                      Tester.Print "rv0="; rv0
                      
                       ' this chip , do not need to test light
                       ' If rv0 <> 0 Then
                       '   If LightON <> &HBF Or LightOFF <> &HFF Then
                       '   Tester.Print "LightON="; LightON
                       '   Tester.Print "LightOFF="; LightOFF
                       '   UsbSpeedTestResult = GPO_FAIL
                       '   rv0 = 3
                       '   End If
                       ' End If
                    
                     
                     
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                 Call LabelMenu(1, rv1, rv0)
            
                '      Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76)
                  
                   
               
                  
                 Call MsecDelay(0.01)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                 
                Call MsecDelay(0.01)
                 
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
              '       Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
                
              
                
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
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
Public Sub AU6433ESF22TestSub()
      
  Tester.Print "AU6433ES : NB mode test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'  this code from AU6433EFTestSub at AU6433EFF21
'  ' add : NB mode test
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
                ChipString = "vid"
                 If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                  Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1#)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
                      
                      Tester.Print "rv0="; rv0
                      
                       ' this chip , do not need to test light
                       ' If rv0 <> 0 Then
                       '   If LightON <> &HBF Or LightOFF <> &HFF Then
                       '   Tester.Print "LightON="; LightON
                       '   Tester.Print "LightOFF="; LightOFF
                       '   UsbSpeedTestResult = GPO_FAIL
                       '   rv0 = 3
                       '   End If
                       ' End If
                    
                     
                     
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                 Call LabelMenu(1, rv1, rv0)
            
                '      Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76)
                  
                   
               
                  
                 Call MsecDelay(0.01)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                 
                Call MsecDelay(0.01)
                 
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
              '       Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
                
              
                
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
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

Public Sub AU6433ESF23TestSub()
      
  Tester.Print "AU6433ES : NB mode test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'  this code from AU6433EFTestSub at AU6433EFF21
'  ' add : NB mode test
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
                ChipString = "vid"
                 If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                  Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1#)
                     
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
                      
                      Tester.Print "rv0="; rv0
                      
                       ' this chip , do not need to test light
                       ' If rv0 <> 0 Then
                       '   If LightON <> &HBF Or LightOFF <> &HFF Then
                       '   Tester.Print "LightON="; LightON
                       '   Tester.Print "LightOFF="; LightOFF
                       '   UsbSpeedTestResult = GPO_FAIL
                       '   rv0 = 3
                       '   End If
                       ' End If
                    
                     
                     
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                 Call LabelMenu(1, rv1, rv0)
            
                '      Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H76)
                  
                   
               
                  
                 Call MsecDelay(0.01)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                 
                Call MsecDelay(0.01)
                 
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
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
              '       Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H57)
                
              
                
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                Call MsecDelay(0.1)
                
                  OpenPipe
                 rv5 = ReInitial(0)
                  
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                If rv5 = 1 Then
                  rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                  
                  If rv5 <> 1 Then
                     rv5 = 2
                     Tester.Print "MS bus width fail"
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
Public Sub AU6433GSTest()
      
  Tester.Print "AU6433EF : NB mode test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'  this code from AU6433EFTestSub at AU6433EFF21
'  ' add : NB mode test
'
'===================================================================


                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
                ChipString = "vid"
                 If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                  Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1#)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
                      
                      Tester.Print "rv0="; rv0
                      
                       ' this chip , do not need to test light
                       ' If rv0 <> 0 Then
                       '   If LightON <> &HBF Or LightOFF <> &HFF Then
                       '   Tester.Print "LightON="; LightON
                       '   Tester.Print "LightOFF="; LightOFF
                       '   UsbSpeedTestResult = GPO_FAIL
                       '   rv0 = 3
                       '   End If
                       ' End If
                    
                     
                     
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = 1  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                 Call LabelMenu(1, rv1, rv0)
            
                '      Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = 1   ' to complete the SMC asbolish
               
              
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
                 
                  
                ReaderExist = 0
                 
                ClosePipe
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  MS Card test
                '================================================
                   
                     
                     rv4 = 1  'AU6344 has no MS slot pin
               
              '       Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
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
                 
                ReaderExist = 0
                
                
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



Public Sub AU6433BSTest()
      
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================

                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
              
                 
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                  Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1#)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
                      
                      Tester.Print "rv0="; rv0
                     
                       
                     
                     
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = 1  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                 Call LabelMenu(1, rv1, rv0)
            
                      Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = 1   ' to complete the SMC asbolish
               
              
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
                 
                  
                ReaderExist = 0
                 
                ClosePipe
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  MS Card test
                '================================================
                   
                     
                     rv4 = 1  'AU6344 has no MS slot pin
               
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
                 
                ReaderExist = 0
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                 
                  CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
                
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
Public Sub AU6433KFF22TestSub()
 Tester.Print "AU6433KF : NB mode"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================

                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
              
                     
                    ChipString = "vid"
                 If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                  Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1#)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
                      
                      Tester.Print "rv0="; rv0
                     
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                          Tester.Print "LightON="; LightOn
                          Tester.Print "LightOFF="; LightOff
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
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
               '  Call LabelMenu(1, rv1, rv0)
            
               '       Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                 rv3 = rv2
                
                '===============================================
                '  MS Card test
                '================================================
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                  '   Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H5E) 'SD + MS
                
                 
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
             
                
                Call MsecDelay(0.1)
                OpenPipe
                rv5 = ReInitial(0)
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                Call LabelMenu(31, rv5, rv0)
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

Public Sub AU6433KFF23TestSub()
 Tester.Print "AU6433KF : NB mode"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================

                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
              
                     
                    ChipString = "vid"
                 If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                  Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1#)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      
                      If rv0 = 1 Then
                         rv0 = Read_SD_Speed(0, 0, 64, "4Bits")
                         If rv0 <> 1 Then
                            rv0 = 2
                            Tester.Print "SD bus width Fail"
                         End If
                      End If
                      
                      
                      
                      ClosePipe
                      
                      Tester.Print "rv0="; rv0
                     
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                          Tester.Print "LightON="; LightOn
                          Tester.Print "LightOFF="; LightOff
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
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
               '  Call LabelMenu(1, rv1, rv0)
            
               '       Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                 rv3 = rv2
                
                '===============================================
                '  MS Card test
                '================================================
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                  '   Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H5E) 'SD + MS
                
                 
                 Call MsecDelay(0.1)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
             
                
                Call MsecDelay(0.1)
                OpenPipe
                rv5 = ReInitial(0)
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                   If rv5 = 1 Then
                         rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                         If rv5 <> 1 Then
                            rv5 = 2
                            Tester.Print "MS bus width Fail"
                         End If
                      End If
                
                Call LabelMenu(31, rv5, rv0)
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

Public Sub AU6433KFTest()
 Tester.Print "AU6433KF : NB mode"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================

                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
              
                     
                    ChipString = "vid"
                 If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                  Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1#)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
                      
                      Tester.Print "rv0="; rv0
                     
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                          Tester.Print "LightON="; LightOn
                          Tester.Print "LightOFF="; LightOff
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
               
                  rv1 = 1  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                 Call LabelMenu(1, rv1, rv0)
            
                      Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = 1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                 rv3 = 1
                
                '===============================================
                '  MS Card test
                '================================================
                   
                     
                     rv4 = 1  'AU6344 has no MS slot pin
               
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
                If rv0 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
             
                  Call MsecDelay(AU6371EL_BootTime * 2)
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect Down Fail"
                    End
                 End If
                 
                ReaderExist = 0
                
             
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                Call LabelMenu(31, rv5, rv0)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                  End If
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

Public Sub AU6433HFTest()
Tester.Print "AU6433HF: Normal mode "
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================

                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
              
                    ChipString = "vid"
                 If GetDeviceName(ChipString) = "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                  Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1#)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
                      
                      Tester.Print "rv0="; rv0
                     
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                          Tester.Print "LightON="; LightOn
                          Tester.Print "LightOFF="; LightOff
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
               
                  rv1 = 1  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                 Call LabelMenu(1, rv1, rv0)
            
                      Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = 1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                 rv3 = 1
                
                '===============================================
                '  MS Card test
                '================================================
                   
                     
                     rv4 = 1  'AU6344 has no MS slot pin
               
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
                If rv0 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
             
             
            
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                Call LabelMenu(31, rv5, rv0)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                  End If
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

Public Sub AU6433FSF28TestSub()
Call PowerSet2(1, "3.3", "0.05", 1, "2.1", "0.05", 1)
Tester.Print "AU6433FS: Normal mode "
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================

                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
              
                    ChipString = "vid"
                 If GetDeviceName(ChipString) = "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                  
             
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
                      
                      Tester.Print "rv0="; rv0
                      
                    
                     
                     
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
               '  Call LabelMenu(1, rv1, rv0)
            
                   '   Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
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
              '  Call MsecDelay(AU6371EL_BootTime)
              '  If CardResult <> 0 Then
              '      MsgBox "Set XD Card Detect Down Fail"
              '      End
              '   End If
                 
                  
               ' ReaderExist = 0
                 
                ClosePipe
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  MS Card test
                '================================================
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect On Fail"
                    End
                 End If
                
                 Call MsecDelay(0.03)
                If rv0 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
             
             
            
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                   If rv5 = 1 Then
                          rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                          If rv5 <> 1 Then
                             rv5 = 2
                             Tester.Print "MS bus width Fail"
                           End If
                      End If
                
                
                Call LabelMenu(31, rv5, rv0)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                  End If
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
Public Sub AU6433FSF29TestSub()


Dim TmpLBA As Long
Dim i As Integer
Dim RT_Flag As String
RT_Flag = ""

RT_Label:

Call PowerSet2(1, "3.3", "0.05", 1, "2.1", "0.05", 1)
Tester.Print "AU6433FS: Normal mode "
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================

                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
'                 CardResult = DO_WritePort(card, Channel_P1A, &H80)
'                    If CardResult <> 0 Then
'                    MsgBox "Power off fail"
'                    End
'                 End If
'                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1.3)    'power on time
              
                    ChipString = "vid"
                 If GetDeviceName(ChipString) = "" Then
                    rv0 = 0
                    
 '                   If RT_Flag = "" Then
 '                       RT_Flag = "Unknow"
 '                       CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
 '                       GoTo RT_Label
 '                   End If
                    GoTo AU6371ELResult
                  
                  End If
                 
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                  
             
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
                      
                      Tester.Print "rv0="; rv0
                      
'=======================================================================================
    'SD R / W
'=======================================================================================
                      
                     TmpLBA = LBA
                     'LBA = 99
                         'For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              
'                                If RT_Flag = "" Then
'                                    RT_Flag = "SD R/W"
'                                    CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
'                                    GoTo RT_Label
'                                End If
                              LBA = TmpLBA
                             GoTo AU6371ELResult
                             End If
                         'Next
                      LBA = TmpLBA
                      
'=======================================================================================
                     
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
               '  Call LabelMenu(1, rv1, rv0)
            
                   '   Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
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
              '  Call MsecDelay(AU6371EL_BootTime)
              '  If CardResult <> 0 Then
              '      MsgBox "Set XD Card Detect Down Fail"
              '      End
              '   End If
                 
                  
               ' ReaderExist = 0
                 
                ClosePipe
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  MS Card test
                '================================================
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect On Fail"
                    End
                 End If
                
                 Call MsecDelay(0.03)
                If rv0 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
             
             
            
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                   If rv5 = 1 Then
                          rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                          If rv5 <> 1 Then
                             rv5 = 2
                             Tester.Print "MS bus width Fail"
                           End If
                      End If
                
                
                Call LabelMenu(31, rv5, rv0)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                  End If
                  CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
'                If (rv5 * rv4 * rv3 * rv2 * rv1 * rv0 <> PASS) And (RT_Flag = "") Then
'                    GoTo RT_Label
'                    RT_Flag = "Retest"
'                End If
                
                 
                
                
AU6371ELResult:
                
                Call PowerSet2(1, "0.0", "0.3", 1, "0.0", "0.3", 1)

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

Public Sub AU6433FSF2ATestSub()


Dim TmpLBA As Long
Dim i As Integer

Tester.Print "AU6433FS: Normal mode "
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'===================================================================

Dim ChipString As String
Dim AU6371EL_SD As Byte
Dim AU6371EL_CF As Byte
Dim AU6371EL_XD As Byte
Dim AU6371EL_MS As Byte
Dim AU6371EL_MSP  As Byte
Dim AU6371EL_BootTime As Single
OldChipName = ""
                 
' initial condition
                
    AU6371EL_SD = 1
    AU6371EL_CF = 2
    AU6371EL_XD = 8
    AU6371EL_MS = 32
    AU6371EL_MSP = 64
    AU6371EL_BootTime = 0.6
                    
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
    ChipString = "vid_058f"
                             
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
                  
    Call MsecDelay(0.1)
    rv0 = WaitDevOn(ChipString)
    Call MsecDelay(0.1)
    
    '===============================================
    '  SD Card test
    '===============================================
    
    ' set SD card detect down
    CardResult = DO_WritePort(card, Channel_P1A, &H7E)
    Call MsecDelay(0.1)
    
    If CardResult <> 0 Then
        MsgBox "Set SD Card Detect Down Fail"
        End
    End If
                     
    
    ClosePipe
    
    If rv0 = 1 Then
        rv0 = CBWTest_New(0, 1, ChipString)
                      
        If rv0 = 1 Then
            rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
            
            If rv0 <> 1 Then
                rv0 = 2
                Tester.Print "SD bus width Fail"
            End If
        End If
        
        ClosePipe
    
    End If
                     
    '=======================================================================================
    'SD R/W 64K
    '=======================================================================================
                      
    If rv0 = 1 Then
        LBA = LBA + 199
        ClosePipe
        rv0 = CBWTest_New_128_Sector_AU6377(0, rv0)  ' write
        If rv0 <> 1 Then
            Tester.Print "R/W 64K Fail ... "
        End If
        ClosePipe
    End If
    
    Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
    Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


    '===============================================
    '  CF Card test
    '================================================
               
    rv1 = rv0  '----------- AU633FS no CF slot
                 
    'Call LabelMenu(1, rv1, rv0)
    'Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    
    '===============================================
    '  SMC Card test  : stop these test for card not enough
    '================================================
              
    'AU6433FS has no SMC slot
               
    rv2 = rv1   ' to complete the SMC asbolish
               
              
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
                     
    rv4 = rv3  'AU6344FS no MS slot
               
    'Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    
    
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
        ClosePipe
        rv5 = CBWTest_New(0, rv4, ChipString)
        If rv5 = 1 Then
            rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
            If rv5 <> 1 Then
                rv5 = 2
                Tester.Print "MS bus width Fail"
            End If
        End If
                
        Call LabelMenu(31, rv5, rv0)
        Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
        ClosePipe
    End If
    
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

Public Sub AU6433FSF0ATestSub()


Dim TmpLBA As Long
Dim i As Integer
Dim HV_Flag As Boolean
Dim HV_Result As String
Dim LV_Result As String

Tester.Print "AU6433FS: Normal mode HV+LV test "

'==================================================================
'
'  this code come from AU6371ELTestSub
'
'===================================================================

Dim ChipString As String
Dim AU6371EL_SD As Byte
Dim AU6371EL_CF As Byte
Dim AU6371EL_XD As Byte
Dim AU6371EL_MS As Byte
Dim AU6371EL_MSP  As Byte
Dim AU6371EL_BootTime As Single
OldChipName = ""
                 
' initial condition
                
    AU6371EL_SD = 1
    AU6371EL_CF = 2
    AU6371EL_XD = 8
    AU6371EL_MS = 32
    AU6371EL_MSP = 64
    AU6371EL_BootTime = 0.6
    
                    
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
    ChipString = "vid_058f"
    HV_Flag = False
    HV_Result = ""
    LV_Result = ""
                             
    '=========================================
    '    POWER on
    '=========================================

Routine_Label:
                
    If (HV_Flag = False) Then
        Call PowerSet2(1, "5.3", "0.2", 1, "5.3", "0.2", 1)
        Tester.Print "Begin HV Test ..."
    Else
        Call PowerSet2(1, "4.7", "0.2", 1, "4.7", "0.2", 1)
        Tester.Print vbCrLf & "Begin LV Test ..."
    End If
    
    CardResult = DO_WritePort(card, Channel_P1A, &H80)
    
    If CardResult <> 0 Then
        MsgBox "Power off fail"
        End
    End If
    
    Call MsecDelay(0.2)
             
    CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
    Call MsecDelay(0.1)
    rv0 = WaitDevOn(ChipString)
    Call MsecDelay(0.1)
    
    '===============================================
    '  SD Card test
    '===============================================
    
    ' set SD card detect down
    CardResult = DO_WritePort(card, Channel_P1A, &H7E)
    Call MsecDelay(0.1)
    
    If CardResult <> 0 Then
        MsgBox "Set SD Card Detect Down Fail"
        End
    End If
                     
    
    ClosePipe
    
    If rv0 = 1 Then
        rv0 = CBWTest_New(0, 1, ChipString)
                      
        If rv0 = 1 Then
            rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
            
            If rv0 <> 1 Then
                rv0 = 2
                Tester.Print "SD bus width Fail"
            End If
        End If
        
        ClosePipe
    
    End If
                     
    '=======================================================================================
    'SD R/W 64K
    '=======================================================================================
                      
    If rv0 = 1 Then
        LBA = LBA + 199
        ClosePipe
        rv0 = CBWTest_New_128_Sector_AU6377(0, rv0)  ' write
        If rv0 <> 1 Then
            Tester.Print "R/W 64K Fail ... "
        End If
        ClosePipe
    End If
    
    Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
    Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


    '===============================================
    '  CF Card test
    '================================================
               
    rv1 = rv0  '----------- AU633FS no CF slot
                 
    'Call LabelMenu(1, rv1, rv0)
    'Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    
    '===============================================
    '  SMC Card test  : stop these test for card not enough
    '================================================
              
    'AU6433FS has no SMC slot
               
    rv2 = rv1   ' to complete the SMC asbolish
               
              
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
                     
    rv4 = rv3  'AU6344FS no MS slot
               
    'Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    
    
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
        ClosePipe
        rv5 = CBWTest_New(0, rv4, ChipString)
        If rv5 = 1 Then
            rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
            If rv5 <> 1 Then
                rv5 = 2
                Tester.Print "MS bus width Fail"
            End If
        End If
                
        Call LabelMenu(31, rv5, rv0)
        Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
        ClosePipe
    End If
    
    
AU6371ELResult:
        
    'Call PowerSet2(1, "0", "0.05", 1, "0", "0.05", 1)       'Purpose to solve over-current on SSOP chip contact moment
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                        
    If HV_Flag = False Then
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
        rv0 = 0
        rv5 = 0
        HV_Flag = True
        Call MsecDelay(0.2)
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
    
End Sub

Public Sub AU6433FSF23TestSub()
Tester.Print "AU6433FS: Normal mode "
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================

                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
              
                    ChipString = "vid"
                 If GetDeviceName(ChipString) = "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                  
             
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
                      
                      Tester.Print "rv0="; rv0
                      
                    
                     
                     
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
               '  Call LabelMenu(1, rv1, rv0)
            
                   '   Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
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
              '  Call MsecDelay(AU6371EL_BootTime)
              '  If CardResult <> 0 Then
              '      MsgBox "Set XD Card Detect Down Fail"
              '      End
              '   End If
                 
                  
               ' ReaderExist = 0
                 
                ClosePipe
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  MS Card test
                '================================================
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect On Fail"
                    End
                 End If
                
                 Call MsecDelay(0.03)
                If rv0 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
             
             
            
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                   If rv5 = 1 Then
                          rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                          If rv5 <> 1 Then
                             rv5 = 2
                             Tester.Print "MS bus width Fail"
                           End If
                      End If
                
                
                Call LabelMenu(31, rv5, rv0)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                  End If
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

Public Sub AU6433FSF22TestSub()
Tester.Print "AU6433FS: Normal mode "
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================

                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
              
                    ChipString = "vid"
                 If GetDeviceName(ChipString) = "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                  
             
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
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
                      
                      Tester.Print "rv0="; rv0
                      
                    
                     
                     
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
               '  Call LabelMenu(1, rv1, rv0)
            
                   '   Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
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
              '  Call MsecDelay(AU6371EL_BootTime)
              '  If CardResult <> 0 Then
              '      MsgBox "Set XD Card Detect Down Fail"
              '      End
              '   End If
                 
                  
               ' ReaderExist = 0
                 
                ClosePipe
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  MS Card test
                '================================================
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect On Fail"
                    End
                 End If
                
                 Call MsecDelay(0.03)
                If rv0 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
             
             
            
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                Call LabelMenu(31, rv5, rv0)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                  End If
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

Public Sub AU6433HSF23TestSub()
Tester.Print "AU6433HS: Normal mode "
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================

                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
              
                    ChipString = "vid"
                 If GetDeviceName(ChipString) = "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 
                 '================================================
          '         CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
                      
                  
          '         If CardResult <> 0 Then
          '          MsgBox "Read light off fail"
          '          End
          '         End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                  
             
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
                     
                  '   CardResult = DO_ReadPort(card, Channel_P1B, LightON)
                  '    If CardResult <> 0 Then
                  '        MsgBox "Read light On fail"
                  '        End
                   '   End If
                     
                     
                             
                           
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
                      
                      Tester.Print "rv0="; rv0
                     
                     '   If rv0 <> 0 Then
                     '     If LightON <> &HBF Or LightOFF <> &HFF Then
                     '     Tester.Print "LightON="; LightON
                     '     Tester.Print "LightOFF="; LightOFF
                     '     UsbSpeedTestResult = GPO_FAIL
                     '     rv0 = 3
                     '     End If
                     '   End If
                    
                     
                     
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
               '  Call LabelMenu(1, rv1, rv0)
            
                   '   Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = 1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                 rv3 = 1
                
                '===============================================
                '  MS Card test
                '================================================
                   
                     
                     rv4 = 1  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect On Fail"
                    End
                 End If
                
                 Call MsecDelay(0.03)
                If rv0 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
             
             
            
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                 If rv5 = 1 Then
                         rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                         If rv5 <> 1 Then
                            rv5 = 2
                            Tester.Print "MS bus width Fail"
                            End If
                      End If
                Call LabelMenu(31, rv5, rv0)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                  End If
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
Public Sub AU6433HSF22TestSub()
Tester.Print "AU6433HS: Normal mode "
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================

                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
              
                    ChipString = "vid"
                 If GetDeviceName(ChipString) = "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 
                 '================================================
          '         CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
                      
                  
          '         If CardResult <> 0 Then
          '          MsgBox "Read light off fail"
          '          End
          '         End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                  
             
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
                     
                  '   CardResult = DO_ReadPort(card, Channel_P1B, LightON)
                  '    If CardResult <> 0 Then
                  '        MsgBox "Read light On fail"
                  '        End
                   '   End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
                      
                      Tester.Print "rv0="; rv0
                     
                     '   If rv0 <> 0 Then
                     '     If LightON <> &HBF Or LightOFF <> &HFF Then
                     '     Tester.Print "LightON="; LightON
                     '     Tester.Print "LightOFF="; LightOFF
                     '     UsbSpeedTestResult = GPO_FAIL
                     '     rv0 = 3
                     '     End If
                     '   End If
                    
                     
                     
                        
                    
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
                '===============================================
                '  CF Card test
                '================================================
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
               '  Call LabelMenu(1, rv1, rv0)
            
                   '   Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = 1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                 rv3 = 1
                
                '===============================================
                '  MS Card test
                '================================================
                   
                     
                     rv4 = 1  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect On Fail"
                    End
                 End If
                
                 Call MsecDelay(0.03)
                If rv0 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
             
             
            
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                Call LabelMenu(31, rv5, rv0)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                  End If
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
Public Sub AU6433HFF22TestSub()
Tester.Print "AU6433HF: Normal mode "
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================

                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
              
                    ChipString = "vid"
                 If GetDeviceName(ChipString) = "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                  
             
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
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
                      
                      Tester.Print "rv0="; rv0
                     
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                          Tester.Print "LightON="; LightOn
                          Tester.Print "LightOFF="; LightOff
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
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
               '  Call LabelMenu(1, rv1, rv0)
            
                   '   Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = 1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                 rv3 = 1
                
                '===============================================
                '  MS Card test
                '================================================
                   
                     
                     rv4 = 1  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect On Fail"
                    End
                 End If
                
                 Call MsecDelay(0.03)
                If rv0 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
             
             
            
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                Call LabelMenu(31, rv5, rv0)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                  End If
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

Public Sub AU6433CSF28TestSub()

Call PowerSet2(1, "3.3", "0.05", 1, "2.1", "0.05", 1)
Tester.Print "AU6433CS: Normal mode "
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================

                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
              
                    ChipString = "vid"
                 If GetDeviceName(ChipString) = "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                  
             
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
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      
                      If rv0 = 1 Then
                       rv0 = Read_SD_Speed(0, 0, 64, "4Bits")
                       If rv0 <> 1 Then
                          rv0 = 2
                          Tester.Print "SD bus width Fail"
                       End If
                     End If
                       
                       
                       
                      ClosePipe
                      
                      Tester.Print "rv0="; rv0
                     
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                          Tester.Print "LightON="; LightOn
                          Tester.Print "LightOFF="; LightOff
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
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
               '  Call LabelMenu(1, rv1, rv0)
            
                   '   Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = 1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                 rv3 = 1
                
                '===============================================
                '  MS Card test
                '================================================
                   
                     
                     rv4 = 1  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect On Fail"
                    End
                 End If
                
                 Call MsecDelay(0.03)
                If rv0 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
             
             
            
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                 If rv5 = 1 Then
                       rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                       If rv5 <> 1 Then
                          rv5 = 2
                          Tester.Print "MS bus width Fail"
                       End If
                     End If
                       
                
                
                Call LabelMenu(31, rv5, rv0)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                  End If
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
Public Sub AU6433CSF29TestSub()

'2011/6/8: Revise wait devon flow, for AU6433S61-RCS FT2 test

Dim TmpLBA As Long
Dim i As Integer

Tester.Print "AU6433CS: Normal mode "

'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================

                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                AU6371EL_SD = 1
                AU6371EL_CF = 2
                AU6371EL_XD = 8
                AU6371EL_MS = 32
                AU6371EL_MSP = 64
                    
                AU6371EL_BootTime = 0.6
                     
                    
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
                
                CardResult = DO_WritePort(card, Channel_P1A, &H0)
                If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                End If
                
                Call MsecDelay(0.02)
                
                CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                End If
                 
                ChipString = "vid"
                
                
                
                '===============================================
                '  SD Card test
                '
                  
             
                '===========================================
                'NO card test
                '============================================
  
                ' set SD card detect down
                CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                Call MsecDelay(0.02)
                Call PowerSet2(1, "3.3", "0.05", 1, "2.1", "0.05", 1)
                
                If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect Down Fail"
                    End
                End If
                
                'Call MsecDelay(0.2)
                rv0 = WaitDevOn(ChipString)
                Call MsecDelay(0.02)
                     
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                
                If CardResult <> 0 Then
                    MsgBox "Read light On fail"
                    End
                End If
                        
                If rv0 = 1 Then
                    If LightOn <> &HBF Or LightOff <> &HFF Then
                        Tester.Print "LightON="; LightOn
                        Tester.Print "LightOFF="; LightOff
                        UsbSpeedTestResult = GPO_FAIL
                        rv0 = 3
                    End If
                End If
                
                ClosePipe
                If rv0 = 1 Then
                    rv0 = CBWTest_New(0, 1, ChipString)
                End If
                
                If rv0 = 1 Then
                    rv0 = Read_SD_Speed(0, 0, 64, "4Bits")
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
                    End If
                End If
                      
                ClosePipe
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      
                TmpLBA = LBA
                rv1 = 0
                LBA = LBA + 199
                            
                rv0 = CBWTest_New_128_Sector_AU6377(0, rv0)  ' write
                 
                If rv0 <> 1 Then
                    LBA = TmpLBA
                    GoTo AU6371ELResult
                End If
                
                LBA = TmpLBA
                      
'=======================================================================================
                     
                Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  CF Card test
                '================================================
               
                rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                'Call LabelMenu(1, rv1, rv0)
            
                '   Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                rv2 = 1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                
                rv3 = 1
                
                '===============================================
                '  MS Card test
                '================================================
                     
                rv4 = 1  'AU6344 has no MS slot pin
               
                ' Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                CardResult = DO_WritePort(card, Channel_P1A, &H5E)  'SD+MS
                Call MsecDelay(0.02)
                CardResult = DO_WritePort(card, Channel_P1A, &H5F)  'MS
                rv5 = ReInitial(0)
                ClosePipe
                
                If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect On Fail"
                    End
                End If
                
                If rv0 = 1 Then
                    rv5 = CBWTest_New(0, rv4, ChipString)
                 
                    If rv5 = 1 Then
                        rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                        If rv5 <> 1 Then
                            rv5 = 2
                            Tester.Print "MS bus width Fail"
                        End If
                    End If
                    
                
                    Call LabelMenu(31, rv5, rv0)
                    Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                    ClosePipe
                End If
                
                
AU6371ELResult:
        
                Call PowerSet2(1, "0", "0.05", 1, "0", "0.05", 1)       'Purpose to solve over-current on SSOP chip contact moment
                CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                        
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

Public Sub AU6433CSF2ATestSub()

'2011/6/8: Revise wait devon flow, for AU6433S61-RCS FT2 test
'2011/7/4: Using internal PWR(V33,V18) bonding issue

Dim TmpLBA As Long
Dim i As Integer

Tester.Print "AU6433CS: Normal mode "

'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================

                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                AU6371EL_SD = 1
                AU6371EL_CF = 2
                AU6371EL_XD = 8
                AU6371EL_MS = 32
                AU6371EL_MSP = 64
                    
                AU6371EL_BootTime = 0.6
                     
                    
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
                
                Call MsecDelay(0.2)
                
                'CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
                'If CardResult <> 0 Then
                '    MsgBox "Read light off fail"
                '    End
                'End If
                 
                ChipString = "vid_058f"
                
                
                
                '===============================================
                '  SD Card test
                '
                  
             
                '===========================================
                'NO card test
                '============================================
  
                ' set SD card detect down
                CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                'Call MsecDelay(0.02)
                'Call PowerSet2(1, "3.3", "0.05", 1, "2.1", "0.05", 1)
                
                If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect Down Fail"
                    End
                End If
                
                Call MsecDelay(0.2)
                rv0 = WaitDevOn(ChipString)
                Call MsecDelay(0.02)
                     
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                
                If CardResult <> 0 Then
                    MsgBox "Read light On fail"
                    End
                End If
                        
                If rv0 = 1 Then
                    'If LightON <> &HBF Or LightOFF <> &HFF Then
                    If LightOn <> &HBF Then
                        Tester.Print "LightON="; LightOn
                        'Tester.Print "LightOFF="; LightOFF
                        UsbSpeedTestResult = GPO_FAIL
                        rv0 = 3
                    End If
                End If
                
                ClosePipe
                If rv0 = 1 Then
                    rv0 = CBWTest_New(0, 1, ChipString)
                End If
                
                If rv0 = 1 Then
                    rv0 = Read_SD_Speed(0, 0, 64, "4Bits")
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
                    End If
                End If
                      
                ClosePipe
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      
                TmpLBA = LBA
                rv1 = 0
                LBA = LBA + 199
                            
                rv0 = CBWTest_New_128_Sector_AU6377(0, rv0)  ' write
                 
                If rv0 <> 1 Then
                    LBA = TmpLBA
                    GoTo AU6371ELResult
                End If
                
                LBA = TmpLBA
                      
'=======================================================================================
                     
                Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  CF Card test
                '================================================
               
                rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                'Call LabelMenu(1, rv1, rv0)
            
                '   Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                rv2 = 1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                
                rv3 = 1
                
                '===============================================
                '  MS Card test
                '================================================
                     
                rv4 = 1  'AU6344 has no MS slot pin
               
                ' Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                CardResult = DO_WritePort(card, Channel_P1A, &H5E)  'SD+MS
                Call MsecDelay(0.02)
                CardResult = DO_WritePort(card, Channel_P1A, &H5F)  'MS
                rv5 = ReInitial(0)
                ClosePipe
                
                If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect On Fail"
                    End
                End If
                
                If rv0 = 1 Then
                    rv5 = CBWTest_New(0, rv4, ChipString)
                 
                    If rv5 = 1 Then
                        rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                        If rv5 <> 1 Then
                            rv5 = 2
                            Tester.Print "MS bus width Fail"
                        End If
                    End If
                    
                
                    Call LabelMenu(31, rv5, rv0)
                    Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                    ClosePipe
                End If
                
                
AU6371ELResult:
        
                'Call PowerSet2(1, "0", "0.05", 1, "0", "0.05", 1)       'Purpose to solve over-current on SSOP chip contact moment
                CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                        
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

Public Sub AU6433CSF09TestSub()

'2011/6/8: Revise wait devon flow, for AU6433S61-RCS FT6(HV+LV) test

Dim TmpLBA As Long
Dim i As Integer
Dim HV_Flag As Boolean
Dim HV_Result As String
Dim LV_Result As String

Tester.Print "AU6433CS: Normal mode HV+LV test "

'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================

                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                HV_Flag = False
                HV_Result = ""
                LV_Result = ""
                
                AU6371EL_SD = 1
                AU6371EL_CF = 2
                AU6371EL_XD = 8
                AU6371EL_MS = 32
                AU6371EL_MSP = 64
                    
                AU6371EL_BootTime = 0.6
                     
                    
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
                
                CardResult = DO_WritePort(card, Channel_P1A, &H0)
                If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                End If
                
                Call MsecDelay(0.02)
                
                CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                End If
                 
                ChipString = "vid"
                
                
                
                '===============================================
                '  SD Card test
                '
                  
             
                '===========================================
                'NO card test
                '============================================
  
                ' set SD card detect down
Routine_Label:
                
                CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                Call MsecDelay(0.02)
                
                If (HV_Flag = False) Then
                    Call PowerSet2(1, "3.6", "0.05", 1, "2.2", "0.05", 1)
                    Tester.Print "Begin HV Test ..."
                Else
                    Call PowerSet2(1, "3.0", "0.05", 1, "1.6", "0.05", 1)
                    Tester.Print vbCrLf & "Begin LV Test ..."
                End If
                
                If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect Down Fail"
                    End
                End If
                
                'Call MsecDelay(0.2)
                rv0 = WaitDevOn(ChipString)
                Call MsecDelay(0.02)
                     
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                
                If CardResult <> 0 Then
                    MsgBox "Read light On fail"
                    End
                End If
                        
                If rv0 = 1 Then
                    If LightOn <> &HBF Or LightOff <> &HFF Then
                        Tester.Print "LightON="; LightOn
                        Tester.Print "LightOFF="; LightOff
                        UsbSpeedTestResult = GPO_FAIL
                        rv0 = 3
                    End If
                End If
                
                ClosePipe
                If rv0 = 1 Then
                    rv0 = CBWTest_New(0, 1, ChipString)
                End If
                
                If rv0 = 1 Then
                    rv0 = Read_SD_Speed(0, 0, 64, "4Bits")
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
                    End If
                End If
                      
                ClosePipe
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      
                TmpLBA = LBA
                rv1 = 0
                LBA = LBA + 199
                            
                rv0 = CBWTest_New_128_Sector_AU6377(0, rv0)  ' write
                 
                If rv0 <> 1 Then
                    LBA = TmpLBA
                    GoTo AU6371ELResult
                End If
                
                LBA = TmpLBA
                      
'=======================================================================================
                     
                Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  CF Card test
                '================================================
               
                rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                'Call LabelMenu(1, rv1, rv0)
            
                '   Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                rv2 = 1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                
                rv3 = 1
                
                '===============================================
                '  MS Card test
                '================================================
                     
                rv4 = 1  'AU6344 has no MS slot pin
               
                ' Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                CardResult = DO_WritePort(card, Channel_P1A, &H5E)  'SD+MS
                Call MsecDelay(0.02)
                CardResult = DO_WritePort(card, Channel_P1A, &H5F)  'MS
                rv5 = ReInitial(0)
                ClosePipe
                
                If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect On Fail"
                    End
                End If
                
                If rv0 = 1 Then
                    rv5 = CBWTest_New(0, rv4, ChipString)
                 
                    If rv5 = 1 Then
                        rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                        If rv5 <> 1 Then
                            rv5 = 2
                            Tester.Print "MS bus width Fail"
                        End If
                    End If
                    
                
                    Call LabelMenu(31, rv5, rv0)
                    Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                    ClosePipe
                End If
                
                
AU6371ELResult:
        
                Call PowerSet2(1, "0", "0.05", 1, "0", "0.05", 1)       'Purpose to solve over-current on SSOP chip contact moment
                CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                        
                If HV_Flag = False Then
                    If rv0 * rv5 = 0 Then
                        HV_Result = "Bin2"
                        Tester.Print "HV Unknow"
                    ElseIf rv0 * rv5 <> 1 Then
                        HV_Result = "Fail"
                        Tester.Print "HV Fail"
                    ElseIf rv0 * rv5 = 1 Then
                        HV_Result = "PASS"
                        Tester.Print "HV PASS"
                    End If
                    rv0 = 0
                    rv5 = 0
                    HV_Flag = True
                    Call MsecDelay(0.2)
                    GoTo Routine_Label
                Else
                    If rv0 * rv5 = 0 Then
                        LV_Result = "Bin2"
                        Tester.Print "LV Unknow"
                    ElseIf rv0 * rv5 <> 1 Then
                        LV_Result = "Fail"
                        Tester.Print "LV Fail"
                    ElseIf rv0 * rv5 = 1 Then
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
        
End Sub

Public Sub AU6433CSF0ATestSub()

'2011/6/8: Revise wait devon flow, for AU6433S61-RCS FT6(HV+LV) test
'2011/7/4: Using internal PWR(V33,V18) bonding issue

Dim TmpLBA As Long
Dim i As Integer
Dim HV_Flag As Boolean
Dim HV_Result As String
Dim LV_Result As String

Tester.Print "AU6433CS: Normal mode HV+LV test "
Call PowerSet2(1, "5.0", "0.2", 1, "5.0", "0.2", 1)

'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================

                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                HV_Flag = False
                HV_Result = ""
                LV_Result = ""
                
                AU6371EL_SD = 1
                AU6371EL_CF = 2
                AU6371EL_XD = 8
                AU6371EL_MS = 32
                AU6371EL_MSP = 64
                    
                AU6371EL_BootTime = 0.6
                     
                    
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
                
                Call MsecDelay(0.2)
                
                'CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
                'If CardResult <> 0 Then
                '    MsgBox "Read light off fail"
                '    End
                'End If
                 
                ChipString = "vid_058f"
                
                
                
                '===============================================
                '  SD Card test
                '
                  
             
                '===========================================
                'NO card test
                '============================================
  
                ' set SD card detect down
Routine_Label:
                
                If (HV_Flag = False) Then
                    Call PowerSet2(1, "5.3", "0.2", 1, "5.3", "0.2", 1)
                    Tester.Print "Begin HV Test ..."
                Else
                    Call PowerSet2(1, "4.7", "0.2", 1, "4.7", "0.2", 1)
                    Tester.Print vbCrLf & "Begin LV Test ..."
                End If
                
                CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                Call MsecDelay(0.02)
                
                
                
                If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect Down Fail"
                    End
                End If
                
                Call MsecDelay(0.2)
                rv0 = WaitDevOn(ChipString)
                Call MsecDelay(0.02)
                     
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                
                If CardResult <> 0 Then
                    MsgBox "Read light On fail"
                    End
                End If
                        
                If rv0 = 1 Then
                    'If LightON <> &HBF Or LightOFF <> &HFF Then
                    If LightOn <> &HBF Then
                        Tester.Print "LightON="; LightOn
                        'Tester.Print "LightOFF="; LightOFF
                        UsbSpeedTestResult = GPO_FAIL
                        rv0 = 3
                    End If
                End If
                
                ClosePipe
                If rv0 = 1 Then
                    rv0 = CBWTest_New(0, 1, ChipString)
                End If
                
                If rv0 = 1 Then
                    rv0 = Read_SD_Speed(0, 0, 64, "4Bits")
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
                    End If
                End If
                      
                ClosePipe
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      
                TmpLBA = LBA
                rv1 = 0
                LBA = LBA + 199
                            
                rv0 = CBWTest_New_128_Sector_AU6377(0, rv0)  ' write
                 
                If rv0 <> 1 Then
                    LBA = TmpLBA
                    GoTo AU6371ELResult
                End If
                
                LBA = TmpLBA
                      
'=======================================================================================
                     
                Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  CF Card test
                '================================================
               
                rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                'Call LabelMenu(1, rv1, rv0)
            
                '   Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                rv2 = 1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                
                rv3 = 1
                
                '===============================================
                '  MS Card test
                '================================================
                     
                rv4 = 1  'AU6344 has no MS slot pin
               
                ' Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                CardResult = DO_WritePort(card, Channel_P1A, &H5E)  'SD+MS
                Call MsecDelay(0.02)
                CardResult = DO_WritePort(card, Channel_P1A, &H5F)  'MS
                rv5 = ReInitial(0)
                ClosePipe
                
                If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect On Fail"
                    End
                End If
                
                If rv0 = 1 Then
                    rv5 = CBWTest_New(0, rv4, ChipString)
                 
                    If rv5 = 1 Then
                        rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                        If rv5 <> 1 Then
                            rv5 = 2
                            Tester.Print "MS bus width Fail"
                        End If
                    End If
                    
                
                    Call LabelMenu(31, rv5, rv0)
                    Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                    ClosePipe
                End If
                
                
AU6371ELResult:
        
                'Call PowerSet2(1, "0", "0.05", 1, "0", "0.05", 1)       'Purpose to solve over-current on SSOP chip contact moment
                CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                        
                If HV_Flag = False Then
                    If rv0 * rv5 = 0 Then
                        HV_Result = "Bin2"
                        Tester.Print "HV Unknow"
                    ElseIf rv0 * rv5 <> 1 Then
                        HV_Result = "Fail"
                        Tester.Print "HV Fail"
                    ElseIf rv0 * rv5 = 1 Then
                        HV_Result = "PASS"
                        Tester.Print "HV PASS"
                    End If
                    rv0 = 0
                    rv5 = 0
                    HV_Flag = True
                    Call MsecDelay(0.2)
                    GoTo Routine_Label
                Else
                    If rv0 * rv5 = 0 Then
                        LV_Result = "Bin2"
                        Tester.Print "LV Unknow"
                    ElseIf rv0 * rv5 <> 1 Then
                        LV_Result = "Fail"
                        Tester.Print "LV Fail"
                    ElseIf rv0 * rv5 = 1 Then
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
        
End Sub

Public Sub AU6433HFF23TestSub()
Tester.Print "AU6433HF: Normal mode "
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================

                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
              
                    ChipString = "vid"
                 If GetDeviceName(ChipString) = "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                  
             
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
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      
                      If rv0 = 1 Then
                       rv0 = Read_SD_Speed(0, 0, 64, "4Bits")
                       If rv0 <> 1 Then
                          rv0 = 2
                          Tester.Print "SD bus width Fail"
                       End If
                     End If
                       
                       
                       
                      ClosePipe
                      
                      Tester.Print "rv0="; rv0
                     
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                          Tester.Print "LightON="; LightOn
                          Tester.Print "LightOFF="; LightOff
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
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
               '  Call LabelMenu(1, rv1, rv0)
            
                   '   Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = 1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                 rv3 = 1
                
                '===============================================
                '  MS Card test
                '================================================
                   
                     
                     rv4 = 1  'AU6344 has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect On Fail"
                    End
                 End If
                
                 Call MsecDelay(0.03)
                If rv0 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
             
             
            
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                 If rv5 = 1 Then
                       rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                       If rv5 <> 1 Then
                          rv5 = 2
                          Tester.Print "MS bus width Fail"
                       End If
                     End If
                       
                
                
                Call LabelMenu(31, rv5, rv0)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                  End If
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
Public Sub AU6433DFF22TestSub()
Tester.Print "AU6433DF: normal test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================

                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
              
                   ChipString = "vid"
                 If GetDeviceName(ChipString) = "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                
             
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
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
                      
                      Tester.Print "rv0="; rv0
                     
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                          Tester.Print "LightON="; LightOn
                          Tester.Print "LightOFF="; LightOff
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
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
              '   Call LabelMenu(1, rv1, rv0)
            
              '        Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
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
              '  Call MsecDelay(AU6371EL_BootTime)
              '  If CardResult <> 0 Then
              '      MsgBox "Set XD Card Detect Down Fail"
              '      End
              '   End If
                 
                  
               ' ReaderExist = 0
                 
                ClosePipe
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  MS Card test
                '================================================
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                     'Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
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
             '     Call MsecDelay(AU6371EL_BootTime * 2)
              '   If CardResult <> 0 Then
             '       MsgBox "Set MSPro Card Detect Down Fail"
             '       End
             '    End If
                 
               ' ReaderExist = 0
                
                
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
Public Sub AU6433DFF23TestSub()
Tester.Print "AU6433DF: normal test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================

                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
              
                   ChipString = "vid"
                 If GetDeviceName(ChipString) = "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                
             
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
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      
                      If rv0 = 1 Then
                          rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                          If rv0 <> 1 Then
                            rv0 = 2
                            Tester.Print "SD bus width fail"
                          End If
                       End If
                          
                          
                      
                      ClosePipe
                      
                      Tester.Print "rv0="; rv0
                     
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                          Tester.Print "LightON="; LightOn
                          Tester.Print "LightOFF="; LightOff
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
               
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
              '   Call LabelMenu(1, rv1, rv0)
            
              '        Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = rv1   ' to complete the SMC asbolish
               
              
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
              '  Call MsecDelay(AU6371EL_BootTime)
              '  If CardResult <> 0 Then
              '      MsgBox "Set XD Card Detect Down Fail"
              '      End
              '   End If
                 
                  
               ' ReaderExist = 0
                 
                ClosePipe
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  MS Card test
                '================================================
                   
                     
                     rv4 = rv3  'AU6344 has no MS slot pin
               
                     'Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
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
             '     Call MsecDelay(AU6371EL_BootTime * 2)
              '   If CardResult <> 0 Then
             '       MsgBox "Set MSPro Card Detect Down Fail"
             '       End
             '    End If
                 
               ' ReaderExist = 0
                
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                If rv5 = 1 Then
                    rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                    If rv5 <> 1 Then
                     rv5 = 2
                    Tester.Print "MS bus width Fail"
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
Public Sub AU6433VFF23TestSub()

Tester.Print "AU6433VF: NBMD Test"
'==================================================================
'
'  this code come from AU6433DFF23TestSub
'
'==================================================================

Dim ChipString As String

Dim AU6371EL_SD As Byte
Dim AU6371EL_CF As Byte
Dim AU6371EL_XD As Byte
Dim AU6371EL_MS As Byte
Dim AU6371EL_MSP  As Byte
Dim AU6371EL_BootTime As Single


OldChipName = ""
' initial condition

AU6371EL_SD = 1
AU6371EL_CF = 2
AU6371EL_XD = 8
AU6371EL_MS = 32
AU6371EL_MSP = 64

AU6371EL_BootTime = 0.6
 
                    
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
'CardResult = DO_WritePort(card, Channel_P1A, &H80)
'If CardResult <> 0 Then
'    MsgBox "Power off fail"
'    End
'End If
'Call MsecDelay(0.1)
'
'CardResult = DO_WritePort(card, Channel_P1A, &H7F)
'
'Call MsecDelay(1#)    'power on time
'
'ChipString = "vid"
'If GetDeviceName(ChipString) = "" Then
'    rv0 = 0
'    GoTo AU6371ELResult
'End If
'
''================================================
'CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
'
'If CardResult <> 0 Then
'    MsgBox "Read light off fail"
'    End
'End If
                   
                 
               
'===============================================
'  SD Card test
'===============================================
       
' set SD card detect down
CardResult = DO_WritePort(card, Channel_P1A, &H7E)

If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If

Call MsecDelay(0.3)
rv0 = WaitDevOn("vid_058f")
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
        Tester.Print "SD bus width fail"
    Else
        rv0 = CBWTest_New_128_Sector_PipeReady(0, rv0)  ' write
        If rv0 <> 1 Then
            rv0 = 2
            Tester.Print "SD R/W 64K Fail"
        End If
    End If
End If


ClosePipe

Call LabelMenu(0, rv0, 1)   ' no card test fail

Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
'===============================================
'  CF Card test
'================================================

rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
'Call LabelMenu(1, rv1, rv0)
'Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
'===============================================
'  SMC Card test  : stop these test for card not enough
'================================================

'AU6433 has no SMC slot

rv2 = rv1   ' to complete the SMC asbolish
               
              
'===============================================
'  XD Card test
'================================================
'CardResult = DO_WritePort(card, Channel_P1A, &H7F)
'
'If CardResult <> 0 Then
'    MsgBox "Set XD Card Detect On Fail"
'    End
'End If


Call MsecDelay(0.01)
If rv2 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &H77)
    Call MsecDelay(0.04)
    rv3 = ReInitial(0)
    Call MsecDelay(0.1)
End If
'  Call MsecDelay(AU6371EL_BootTime)
'  If CardResult <> 0 Then
'      MsgBox "Set XD Card Detect Down Fail"
'      End
'   End If

ClosePipe
rv3 = CBWTest_New(0, rv2, ChipString)
ClosePipe
Call LabelMenu(2, rv3, rv2)

Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                
'===============================================
'  MS Card test
'================================================

rv4 = rv3  'AU6344 has no MS slot pin

'Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'===============================================
'  MS Pro Card test
'================================================

'CardResult = DO_WritePort(card, Channel_P1A, &H7F)
'
'If CardResult <> 0 Then
'    MsgBox "Set MSPro Card Detect On Fail"
'    End
'End If

'Call MsecDelay(0.03)

If rv4 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
    Call MsecDelay(0.04)
    rv5 = ReInitial(0)
    Call MsecDelay(0.1)
End If

'If CardResult <> 0 Then
'   MsgBox "Set MSPro Card Detect Down Fail"
'   End
'End If

'ReaderExist = 0

ClosePipe
rv5 = CBWTest_New(0, rv4, ChipString)
If rv5 = 1 Then
    rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
    If rv5 <> 1 Then
        rv5 = 2
        Tester.Print "MS bus width Fail"
    End If
End If

Call LabelMenu(31, rv5, rv4)
Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
ClosePipe
                 
CardResult = DO_WritePort(card, Channel_P1A, &H7F)   ' Close power

If rv5 = 1 Then
    rv5 = WaitDevOFF("vid_058f")
    Call MsecDelay(0.2)
    If rv5 <> 1 Then
        rv5 = 3
        Tester.Print "NBMD Test Fail ..."
    End If
End If

CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
If CardResult <> 0 Then
    MsgBox "Read light off fail"
    End
End If

If rv0 = 1 Then
    If LightOn <> &HBF Or LightOff <> &HFF Then
        Tester.Print "LightON="; LightOn
        Tester.Print "LightOFF="; LightOff
        UsbSpeedTestResult = GPO_FAIL
        rv0 = 3
    End If
End If

                
AU6371ELResult:

CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
    
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
Public Sub AU6433DFTest()
Tester.Print "AU6433DF: normal test"
'==================================================================
'
'  this code come from AU6371ELTestSub
'
'
'===================================================================

                Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
               
                 
                 
                ' initial condition
                
                    AU6371EL_SD = 1
                    AU6371EL_CF = 2
                    AU6371EL_XD = 8
                    AU6371EL_MS = 32
                    AU6371EL_MSP = 64
                    
                     AU6371EL_BootTime = 0.6
                     
                    
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
                 Call MsecDelay(0.1)
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(1#)    'power on time
              
                   ChipString = "vid"
                 If GetDeviceName(ChipString) = "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 '================================================
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                   
                 
               
                '===============================================
                '  SD Card test
                '
                  Call MsecDelay(0.3)
             
                 '===========================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(1#)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
                      
                      Tester.Print "rv0="; rv0
                     
                        If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                          Tester.Print "LightON="; LightOn
                          Tester.Print "LightOFF="; LightOff
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
               
                  rv1 = 1  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                 Call LabelMenu(1, rv1, rv0)
            
                      Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                   rv2 = 1   ' to complete the SMC asbolish
               
              
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
              '  Call MsecDelay(AU6371EL_BootTime)
              '  If CardResult <> 0 Then
              '      MsgBox "Set XD Card Detect Down Fail"
              '      End
              '   End If
                 
                  
               ' ReaderExist = 0
                 
                ClosePipe
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  MS Card test
                '================================================
                   
                     
                     rv4 = 1  'AU6344 has no MS slot pin
               
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
             '     Call MsecDelay(AU6371EL_BootTime * 2)
              '   If CardResult <> 0 Then
             '       MsgBox "Set MSPro Card Detect Down Fail"
             '       End
             '    End If
                 
               ' ReaderExist = 0
                
                
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
