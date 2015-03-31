Attribute VB_Name = "AU6438MDL"
Public AU6438SortingFlag As Boolean

Public Sub AU6438MPTest()

'If ChipName = "AU6438BSF02" Then
'
'    Call AU6438BSF02TestSub
'
'ElseIf ChipName = "AU6438BSF00" Then
'
'    Call AU6438BSF00TestSub
'
'ElseIf ChipName = "AU6438BSF01" Then
'
'    Call AU6438BSF01TestSub
'
'ElseIf ChipName = "AU6438BSF21" Then
'
'    Call AU6438BSF21TestSub
'
'ElseIf ChipName = "AU6438BSF20" Then
'
'    Call AU6438BSF20TestSub
'
'ElseIf ChipName = "AU6438CFF20" Then
'
'    Call AU6438CFF20TestSub
'
'ElseIf ChipName = "AU6438GFF20" Then
'
'    Call AU6438GFF20TestSub
'
'ElseIf ChipName = "AU6438BLF20" Then
'
'    Call AU6438BLF20TestSub
'
'ElseIf ChipName = "AU6438CLF20" Then
'
'    Call AU6438CLF20TestSub
'
'ElseIf ChipName = "AU6438IFS20" Then
'
'    Call AU6438IFS20TestSub
'
'ElseIf ChipName = "AU6438CFF00" Then
'
'    Call AU6438CFF00TestSub
'
'ElseIf ChipName = "AU6438CFS10" Then
'
'    Call AU6438CFS10TestSub
'
'ElseIf ChipName = "AU6438KFE10" Then
'
'    Call AU6438KFE10TestSub
'
'ElseIf ChipName = "AU6438KFS10" Then
'
'    Call AU6438KFS10TestSub


AU6438SortingFlag = False

'Sorting Flow
'BS bonding type only Sorting Flow

If ChipName = "AU6438BSS21" Then
    AU6438SortingFlag = True
    Call AU6438BSS21TestSub
    
ElseIf ChipName = "AU6438BSS01" Then
    AU6438SortingFlag = True
    Call AU6438BSS01TestSub
    
ElseIf ChipName = "AU6438CFS21" Then
    AU6438SortingFlag = True
    Call AU6438CFS21TestSub
  
ElseIf ChipName = "AU6438EFS20" Then
    AU6438SortingFlag = True
    Call AU6438EFS20TestSub

ElseIf ChipName = "AU6438MFS20" Then
    AU6438SortingFlag = True
    Call AU6438MFS20TestSub
    
ElseIf ChipName = "AU6438IFS30" Then
    AU6438SortingFlag = True
    Call AU6438IFS30TestSub

ElseIf ChipName = "AU6438IFS00" Then
    AU6438SortingFlag = True
    Call AU6438IFS00TestSub

ElseIf ChipName = "AU6438KFS30" Then
    AU6438SortingFlag = True
    Call AU6438KFS30TestSub

ElseIf ChipName = "AU6438KFS00" Then
    AU6438SortingFlag = True
    Call AU6438KFS00TestSub

End If


'FT2 flow
If ChipName = "AU6438BSF21" Then
    AU6438SortingFlag = False
    Call AU6438BSS21TestSub

ElseIf ChipName = "AU6438BSF01" Then
    AU6438SortingFlag = False
    Call AU6438BSS01TestSub
    
ElseIf ChipName = "AU6438CFF21" Then
    AU6438SortingFlag = False
    Call AU6438CFS21TestSub

ElseIf ChipName = "AU6438CFF01" Then
    AU6438SortingFlag = False
    Call AU6438CFF01TestSub
  
ElseIf ChipName = "AU6438EFF20" Then
    AU6438SortingFlag = False
    Call AU6438EFS20TestSub

ElseIf ChipName = "AU6438MFF20" Then
    AU6438SortingFlag = False
    Call AU6438MFS20TestSub
    
ElseIf ChipName = "AU6438IFF30" Then
    AU6438SortingFlag = False
    Call AU6438IFS30TestSub

ElseIf ChipName = "AU6438IFF00" Then
    AU6438SortingFlag = False
    Call AU6438IFS00TestSub

ElseIf ChipName = "AU6438KFF30" Then
    AU6438SortingFlag = False
    Call AU6438KFS30TestSub

ElseIf ChipName = "AU6438KFF00" Then
    AU6438SortingFlag = False
    Call AU6438KFS00TestSub

End If


End Sub
Public Sub AU6438BSF20TestSub()
'Tester.Print "AU6433HF: Normal mode "
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
                    ChipString = "vid_058f"
                       
               
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7D)
                  
                 Call MsecDelay(0.1)     'power on time
             
              
                 
                 '================================================
                  
              
                   
                 
               
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
                     Call MsecDelay(0.4)
                     
                    
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
                       
                        CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                       
                      ClosePipe
                      
                      Tester.Print "rv0="; rv0
                     
                         If rv0 <> 0 Then
                           If CAndValue(LightOn, &H2) = 2 Then
                           Tester.Print "LightON="; LightOn
                 
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
                  CardResult = DO_WritePort(card, Channel_P1A, &H7D)   ' Close power
                    Call MsecDelay(0.4)
             
                 If GetDeviceName(ChipString) = "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                
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

Public Sub AU6438BSF21TestSub()


'2011/11/15 purpose to solve AU6438R62-RIF can't measuring GPON7 value
Dim RT_Counter As Integer


Tester.Print "AU6438BSF: Normal mode "
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
                RT_Counter = 0
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
                       
               
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H7D)
                  
                Call MsecDelay(0.1)     'power on time
             
                '===============================================
                '  SD Card test
                '============================================
  
                ' set SD card detect down
                
                CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect Down Fail"
                    End
                End If
                
                rv0 = WaitDevOn(ChipString)
                'Call MsecDelay(0.4)
                
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
                End If
                
                ClosePipe
                      
                Tester.Print "rv0="; rv0
                    
                If rv0 = 1 Then
                    Do
                        Call MsecDelay(0.2)
                        CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                     
                        If rv0 <> 0 Then
                            'If CAndValue(LightON, &H2) = 2 Then
                            If (LightOn <> 252) Then
                                If RT_Counter >= 4 Then
                                    Tester.Print "LightON="; LightOn
                                    UsbSpeedTestResult = GPO_FAIL
                                End If
                                rv0 = 3
                            Else
                                rv0 = 1
                            End If
                        
                        End If
                        
                        RT_Counter = RT_Counter + 1
                    
                    Loop While ((rv0 <> 1) And (RT_Counter < 5))
                 
                End If
                 
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
               
                rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                 rv3 = rv2
                
                
                '===============================================
                '  MS Card test
                '================================================
                rv4 = 1  'AU6344 has no MS slot pin
               
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
                
                If rv0 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
             
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
                
                'CardResult = DO_WritePort(card, Channel_P1A, &H7D)   ' Close power
                'Call MsecDelay(0.4)
             
                'If GetDeviceName(ChipString) <> "" Then
                '    rv0 = 3
                '    Tester.Print "NBMD Test Fail ..."
                'End If
                 
                
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

Public Sub AU6438BSS21TestSub()


'2011/11/15 purpose to solve AU6438R62-RIF can't measuring GPON7 value
Dim RT_Counter As Integer

Tester.Print "AU6438BSS: Normal mode "
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
                RT_Counter = 0
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
                       
                If AU6438SortingFlag Then
                    Call PowerSet2(1, "4.4", "0.3", 1, "4.4", "0.3", 1)
                End If
                
                CardResult = DO_WritePort(card, Channel_P1A, &H7D)
                  
                Call MsecDelay(0.1)     'power on time
             
                '===============================================
                '  SD Card test
                '============================================
  
                ' set SD card detect down
                
                CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
                If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect Down Fail"
                    End
                End If
                
                rv0 = WaitDevOn(ChipString)
                'Call MsecDelay(0.4)
                
                ClosePipe
                
                If rv0 = 1 Then
                    rv0 = CBWTest_New(0, 1, ChipString)
                      
                      
                    If rv0 = 1 Then
                        rv0 = CBWTest_New_128_Sector_PipeReady(0, rv0)    ' write
                    End If
                    
                    If rv0 = 1 Then
                        rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                        If rv0 <> 1 Then
                            rv0 = 2
                            Tester.Print "SD bus width Fail"
                        End If
                     End If
                End If
                
                ClosePipe
                      
                Tester.Print "rv0="; rv0
                    
                If rv0 = 1 Then
                    Do
                        Call MsecDelay(0.2)
                        CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                     
                        If rv0 <> 0 Then
                            'If CAndValue(LightON, &H2) = 2 Then
                            If (LightOn <> 252) Then
                                If RT_Counter >= 4 Then
                                    Tester.Print "LightON="; LightOn
                                    UsbSpeedTestResult = GPO_FAIL
                                End If
                                rv0 = 3
                            Else
                                rv0 = 1
                            End If
                        
                        End If
                        
                        RT_Counter = RT_Counter + 1
                    
                    Loop While ((rv0 <> 1) And (RT_Counter < 5))
                 
                End If
                
                If rv0 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H7D)
                    Call MsecDelay(0.04)
                    CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                    
                    rv0 = CBWTest_New(0, 1, ChipString)
                      
                      
                    If rv0 = 1 Then
                        rv0 = CBWTest_New_128_Sector_PipeReady(0, rv0)    ' write
                    End If
                    ClosePipe
                End If
                
                If rv0 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H7D)
                    Call MsecDelay(0.04)
                    CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                    
                    rv0 = CBWTest_New(0, 1, ChipString)
                      
                      
                    If rv0 = 1 Then
                        rv0 = CBWTest_New_128_Sector_PipeReady(0, rv0)    ' write
                    End If
                    ClosePipe
                End If

              
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
               
                rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                 rv3 = rv2
                
                
                '===============================================
                '  MS Card test
                '================================================
                rv4 = 1  'AU6344 has no MS slot pin
               
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
                
                If rv0 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
             
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
                
                'CardResult = DO_WritePort(card, Channel_P1A, &H7D)   ' Close power
                'Call MsecDelay(0.4)
             
                'If GetDeviceName(ChipString) <> "" Then
                '    rv0 = 3
                '    Tester.Print "NBMD Test Fail ..."
                'End If
                 
                
AU6371ELResult:
                
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
Public Sub AU6438BSF00TestSub()

'2011/5/12: External 3.3V HV(3.6)+LV(3.0) test

            Dim ChipString As String
            Dim AU6371EL_SD As Byte
            Dim AU6371EL_CF As Byte
            Dim AU6371EL_XD As Byte
            Dim AU6371EL_MS As Byte
            Dim AU6371EL_MSP  As Byte
            Dim AU6371EL_BootTime As Single
            Dim HV_Flag As Boolean
            Dim HV_Result As Byte
            Dim LV_Result As Byte
            OldChipName = ""
               
            ' initial condition
                
            AU6371EL_SD = 1
            AU6371EL_CF = 2
            AU6371EL_XD = 8
            AU6371EL_MS = 32
            AU6371EL_MSP = 64
            AU6371EL_BootTime = 0.6
                     
            HV_Flag = False
            HV_Result = 0
            LV_Result = 0
                    
            If PCI7248InitFinish = 0 Then
                PCI7248Exist
            End If
               
            ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            ' CardResult = DO_WritePort(card, Channel_P1B, &H0)
ReTest_LABEL:

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
                       

                
                If HV_Flag = False Then
                    Call PowerSet2(1, "3.6", "0.05", 1, "3.6", "0.05", 1)
                    Tester.Print "Start HV(3.6) Test ..."
                Else
                    Tester.Print vbCrLf & vbCrLf & "Start LV(3.0) Test ..."
                    Call PowerSet2(1, "3.0", "0.05", 1, "3.0", "0.05", 1)
                End If
                 
                 
                 
                 'CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 'Call MsecDelay(0.1)     'power on time
             
                 '================================================
               
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
                    
                    ReaderExist = 0
                    
                    Call MsecDelay(0.2)
                    rv0 = WaitDevOn(ChipString)
                    Call MsecDelay(0.02)
                           
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
                    
                    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                    If CardResult <> 0 Then
                        MsgBox "Read light On fail"
                        End
                    End If
                    ClosePipe
                      
                    Tester.Print "rv0="; rv0
                     
                    If rv0 = 1 Then
                        If CAndValue(LightOn, &H2) = 2 Then
                            Tester.Print "LightON="; LightOn
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
              
                CardResult = DO_WritePort(card, Channel_P1A, &H5E)
                '
                'If CardResult <> 0 Then
                '    MsgBox "Set MSPro Card Detect On Fail"
                '    End
                'End If
                
                Call MsecDelay(0.02)
                
                If rv0 = 1 Then
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
                       
                End If
                
                Call LabelMenu(31, rv5, rv0)
                Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                 
                If HV_Flag = False Then
                    HV_Flag = True
                    Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
                    'CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                    Call MsecDelay(0.1)
                    HV_Result = rv0 * rv5
                    rv0 = 0
                    rv5 = 0
                    GoTo ReTest_LABEL
                ElseIf HV_Flag = True Then
                    LV_Result = rv0 * rv5
                    Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
                     CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                    Call MsecDelay(0.1)
                End If
                
AU6371ELResult:
                
               
                
                
                If (HV_Result = 0) And (LV_Result = 0) Then
                    UnknowDeviceFail = UnknowDeviceFail + 1
                    TestResult = "Bin2"
                ElseIf (HV_Result <> 1) And (LV_Result = 1) Then
                    TestResult = "Bin3"
                ElseIf (HV_Result = 1) And (LV_Result <> 1) Then
                    TestResult = "Bin4"
                ElseIf (HV_Result <> 1) And (LV_Result <> 1) Then
                    TestResult = "Bin5"
                ElseIf (HV_Result = 1) And (LV_Result = 1) Then
                    TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                End If
                      
                      
                      
                '      If rv0 = UNKNOW Then
                '           UnknowDeviceFail = UnknowDeviceFail + 1
                '           TestResult = "UNKNOW"
                '        ElseIf rv0 = WRITE_FAIL Then
                '            SDWriteFail = SDWriteFail + 1
               '             TestResult = "SD_WF"
                '        ElseIf rv0 = READ_FAIL Then
                '            SDReadFail = SDReadFail + 1
                '            TestResult = "SD_RF"
                '        ElseIf rv1 = WRITE_FAIL Then
                ''            CFWriteFail = CFWriteFail + 1
                '            TestResult = "CF_WF"
                '        ElseIf rv1 = READ_FAIL Then
                '            CFReadFail = CFReadFail + 1
                '            TestResult = "CF_RF"
                '        ElseIf rv2 = WRITE_FAIL Or rv3 = WRITE_FAIL Then
                 '           XDWriteFail = XDWriteFail + 1
                 '           TestResult = "XD_WF"
                 '       ElseIf rv2 = READ_FAIL Or rv3 = READ_FAIL Then
                 '           XDReadFail = XDReadFail + 1
                 '           TestResult = "XD_RF"
                 '        ElseIf rv4 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                 '           MSWriteFail = MSWriteFail + 1
                  '          TestResult = "MS_WF"
                 '       ElseIf rv4 = READ_FAIL Or rv5 = READ_FAIL Then
                 '           MSReadFail = MSReadFail + 1
                 '           TestResult = "MS_RF"
                 '
                 '
                 '       ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                 '            TestResult = "PASS"
                 '       Else
                 '           TestResult = "Bin2"
                 '
                 '       End If
End Sub
Public Sub AU6438BSF01TestSub()

'2011/9/21: External 5V HV(5.3)+LV(4.7) test

Dim ChipString As String
Dim AU6371EL_SD As Byte
Dim AU6371EL_CF As Byte
Dim AU6371EL_XD As Byte
Dim AU6371EL_MS As Byte
Dim AU6371EL_MSP  As Byte
Dim AU6371EL_BootTime As Single
Dim HV_Flag As Boolean
Dim HV_Result As Byte
Dim LV_Result As Byte
            
    OldChipName = ""
               
    ' initial condition
    AU6371EL_SD = 1
    AU6371EL_CF = 2
    AU6371EL_XD = 8
    AU6371EL_MS = 32
    AU6371EL_MSP = 64
    AU6371EL_BootTime = 0.6
    
    HV_Flag = False
    HV_Result = 0
    LV_Result = 0
                    
    If PCI7248InitFinish = 0 Then
        PCI7248Exist
    End If
               
    ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
    ' CardResult = DO_WritePort(card, Channel_P1B, &H0)

ReTest_LABEL:

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
      
    If HV_Flag = False Then
        Call PowerSet2(1, "5.3", "0.3", 1, "5.3", "0.3", 1)
        Tester.Print "Start HV(5.3) Test ..."
    Else
        Tester.Print vbCrLf & vbCrLf & "Start LV(4.7) Test ..."
        Call PowerSet2(1, "4.7", "0.3", 1, "4.7", "0.3", 1)
    End If
                
    CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
    Call MsecDelay(0.2)     'power on time
             
               
    '===============================================
    '  SD Card test
    '===============================================
                  
    ' set SD card detect down
    CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
    If CardResult <> 0 Then
        MsgBox "Set SD Card Detect Down Fail"
        End
    End If
    
    ReaderExist = 0
                    
    Call MsecDelay(0.2)
    rv0 = WaitDevOn(ChipString)
    Call MsecDelay(0.02)
    ClosePipe
                      
    If rv0 = 1 Then
        rv0 = CBWTest_New(0, rv0, ChipString)
    End If
                    
    If rv0 = 1 Then
        rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
        If rv0 <> 1 Then
            rv0 = 2
            Tester.Print "SD bus width Fail"
        End If
    End If
                    
    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
    If CardResult <> 0 Then
        MsgBox "Read light On fail"
        End
    End If
    ClosePipe
                      
    Tester.Print "rv0="; rv0
                     
    If rv0 = 1 Then
        If CAndValue(LightOn, &H2) = 2 Then
            Tester.Print "LightON="; LightOn
            UsbSpeedTestResult = GPO_FAIL
            rv0 = 3
        End If
    End If
                    
    Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
    Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    '===============================================
    '  CF Card test
    '================================================
               
    rv1 = rv0       '----------- AU6438BS no CF slot
                 
    'Call LabelMenu(1, rv1, rv0)
    'Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    
    '===============================================
    '  SMC Card test  : stop these test for card not enough
    '================================================
                
    rv2 = rv1       '----------- AU6438BS no SMC slot
               
              
    '===============================================
    '  XD Card test
    '================================================
                 
    rv3 = rv2       '----------- AU6438BS no XD slot
                
    
    '===============================================
    '  MS Card test
    '================================================
                     
    rv4 = rv3       '----------- AU6438BS no MS slot
               
    'Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    
    
    '===============================================
    'MS Pro Card test
    '================================================
              
    CardResult = DO_WritePort(card, Channel_P1A, &H5E)
                
    If CardResult <> 0 Then
        MsgBox "Set MSPro Card Detect On Fail"
        End
    End If
                
    Call MsecDelay(0.02)
                
    If rv0 = 1 Then
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
                       
    End If
                
    Call LabelMenu(31, rv5, rv0)
    Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    ClosePipe
                 
    If HV_Flag = False Then
        HV_Flag = True
        Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
        'CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
        Call MsecDelay(0.3)
        HV_Result = rv0 * rv5
        rv0 = 0
        rv5 = 0
        GoTo ReTest_LABEL
    ElseIf HV_Flag = True Then
        LV_Result = rv0 * rv5
        Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
        CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
        Call MsecDelay(0.1)
    End If
                
                
AU6371ELResult:
                
                
    If (HV_Result = 0) And (LV_Result = 0) Then
        UnknowDeviceFail = UnknowDeviceFail + 1
        TestResult = "Bin2"
    ElseIf (HV_Result <> 1) And (LV_Result = 1) Then
        TestResult = "Bin3"
    ElseIf (HV_Result = 1) And (LV_Result <> 1) Then
        TestResult = "Bin4"
    ElseIf (HV_Result <> 1) And (LV_Result <> 1) Then
        TestResult = "Bin5"
    ElseIf (HV_Result = 1) And (LV_Result = 1) Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
                
                
End Sub

Public Sub AU6438BSF02TestSub()

'2011/9/21: External 5V HV(5.3)+LV(4.7) test
'2011/11/15: purpose to solve GPON lightOn delay issue


Dim ChipString As String
Dim AU6371EL_SD As Byte
Dim AU6371EL_CF As Byte
Dim AU6371EL_XD As Byte
Dim AU6371EL_MS As Byte
Dim AU6371EL_MSP  As Byte
Dim AU6371EL_BootTime As Single
Dim HV_Flag As Boolean
Dim HV_Result As Byte
Dim LV_Result As Byte
Dim RT_Counter As Integer

    OldChipName = ""
               
    ' initial condition
    AU6371EL_SD = 1
    AU6371EL_CF = 2
    AU6371EL_XD = 8
    AU6371EL_MS = 32
    AU6371EL_MSP = 64
    AU6371EL_BootTime = 0.6
    
    HV_Flag = False
    HV_Result = 0
    LV_Result = 0
    
    If PCI7248InitFinish = 0 Then
        PCI7248Exist
    End If
               
    ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
    ' CardResult = DO_WritePort(card, Channel_P1B, &H0)

ReTest_LABEL:

    LBA = LBA + 1
                         
    rv0 = 0
    rv1 = 0
    rv2 = 0
    rv3 = 0
    rv4 = 0
    rv5 = 0
    rv6 = 0
    rv7 = 0
    RT_Counter = 0
             
             
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
      
    If HV_Flag = False Then
        Call PowerSet2(1, "5.3", "0.3", 1, "5.3", "0.3", 1)
        Tester.Print "Start HV(5.3) Test ..."
    Else
        Tester.Print vbCrLf & vbCrLf & "Start LV(4.7) Test ..."
        Call PowerSet2(1, "4.7", "0.3", 1, "4.7", "0.3", 1)
    End If
                
    CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
    Call MsecDelay(0.2)     'power on time
             
               
    '===============================================
    '  SD Card test
    '===============================================
                  
    ' set SD card detect down
    CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
    If CardResult <> 0 Then
        MsgBox "Set SD Card Detect Down Fail"
        End
    End If
    
    ReaderExist = 0
                    
    Call MsecDelay(0.2)
    rv0 = WaitDevOn(ChipString)
    Call MsecDelay(0.02)
    ClosePipe
                      
    If rv0 = 1 Then
        rv0 = CBWTest_New(0, rv0, ChipString)
    End If
                    
    If rv0 = 1 Then
        rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
        If rv0 <> 1 Then
            rv0 = 2
            Tester.Print "SD bus width Fail"
        End If
    End If
                    
    ClosePipe
                      
    Tester.Print "rv0="; rv0
                     
    If rv0 = 1 Then
        Do
            Call MsecDelay(0.2)
            CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
            If CardResult <> 0 Then
                MsgBox "Read light On fail"
                End
            End If
    
            'If CAndValue(LightON, &H2) = 2 Then
            If (LightOn <> 252) Then
                If RT_Counter >= 4 Then
                    Tester.Print "LightON="; LightOn
                    UsbSpeedTestResult = GPO_FAIL
                End If
                
                rv0 = 3
            Else
                rv0 = 1
            End If
            
            RT_Counter = RT_Counter + 1
            
        Loop While (rv0 <> 1) And (RT_Counter < 5)
    End If
                    
    Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
    Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    '===============================================
    '  CF Card test
    '================================================
               
    rv1 = rv0       '----------- AU6438BS no CF slot
                 
    'Call LabelMenu(1, rv1, rv0)
    'Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    
    '===============================================
    '  SMC Card test  : stop these test for card not enough
    '================================================
                
    rv2 = rv1       '----------- AU6438BS no SMC slot
               
              
    '===============================================
    '  XD Card test
    '================================================
                 
    rv3 = rv2       '----------- AU6438BS no XD slot
                
    
    '===============================================
    '  MS Card test
    '================================================
                     
    rv4 = rv3       '----------- AU6438BS no MS slot
               
    'Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    
    
    '===============================================
    'MS Pro Card test
    '================================================
              
    CardResult = DO_WritePort(card, Channel_P1A, &H5E)
                
    If CardResult <> 0 Then
        MsgBox "Set MSPro Card Detect On Fail"
        End
    End If
                
    Call MsecDelay(0.02)
                
    If rv0 = 1 Then
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
                       
    End If
                
    Call LabelMenu(31, rv5, rv0)
    Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    ClosePipe
                 
    If HV_Flag = False Then
        HV_Flag = True
        Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
        'CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
        Call MsecDelay(0.3)
        HV_Result = rv0 * rv5
        rv0 = 0
        rv5 = 0
        GoTo ReTest_LABEL
    ElseIf HV_Flag = True Then
        LV_Result = rv0 * rv5
        Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
        CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
        Call MsecDelay(0.1)
    End If
                
                
AU6371ELResult:
                
                
    If (HV_Result = 0) And (LV_Result = 0) Then
        UnknowDeviceFail = UnknowDeviceFail + 1
        TestResult = "Bin2"
    ElseIf (HV_Result <> 1) And (LV_Result = 1) Then
        TestResult = "Bin3"
    ElseIf (HV_Result = 1) And (LV_Result <> 1) Then
        TestResult = "Bin4"
    ElseIf (HV_Result <> 1) And (LV_Result <> 1) Then
        TestResult = "Bin5"
    ElseIf (HV_Result = 1) And (LV_Result = 1) Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
                
                
End Sub

Public Sub AU6438BSS01TestSub()

'2011/9/21: External 5V HV(5.3)+LV(4.7) test
'2011/11/15: purpose to solve GPON lightOn delay issue
'2014/3/18: Using 5.25/4.4 & DP,DM 36ohm condition


Dim ChipString As String
Dim AU6371EL_SD As Byte
Dim AU6371EL_CF As Byte
Dim AU6371EL_XD As Byte
Dim AU6371EL_MS As Byte
Dim AU6371EL_MSP  As Byte
Dim AU6371EL_BootTime As Single
Dim HV_Done_Flag As Boolean
Dim HV_Result As String
Dim LV_Result As String
Dim RT_Counter As Integer

    OldChipName = ""
               
    ' initial condition
    AU6371EL_SD = 1
    AU6371EL_CF = 2
    AU6371EL_XD = 8
    AU6371EL_MS = 32
    AU6371EL_MSP = 64
    AU6371EL_BootTime = 0.6
    
    HV_Done_Flag = False
    HV_Result = 0
    LV_Result = 0
    
    If PCI7248InitFinish_Sync = 0 Then
        PCI7248Exist_P1C_Sync
    End If
               

Routine_Label:

    LBA = LBA + 1
                         
    rv0 = 0
    rv1 = 0
    rv2 = 0
    rv3 = 0
    rv4 = 0
    rv5 = 0
    rv6 = 0
    rv7 = 0
    RT_Counter = 0
             
             
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
    
    If Not HV_Done_Flag Then
        Call PowerSet2(0, "5.25", "0.3", 1, "5.25", "0.3", 1)
        Call MsecDelay(0.2)
        Tester.Print "AU6438BS : 5.25V Begin Test ..."
        SetSiteStatus (RunHV)
    Else
        Call PowerSet2(0, "4.4", "0.3", 1, "4.4", "0.3", 1)
        Call MsecDelay(0.2)
        Tester.Print vbCrLf & "AU6438BS : 4.4V Begin Test ..."
        SetSiteStatus (RunLV)
    End If

    Call MsecDelay(0.2)     'power on time
             
               
    '===============================================
    '  SD Card test
    '===============================================
                  
    ' set SD card detect down
    CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                      
    If CardResult <> 0 Then
        MsgBox "Set SD Card Detect Down Fail"
        End
    End If
    
    ReaderExist = 0
                    
    Call MsecDelay(0.2)
    rv0 = WaitDevOn(ChipString)
    Call MsecDelay(0.02)
    ClosePipe
                      
    If rv0 = 1 Then
        rv0 = CBWTest_New(0, rv0, ChipString)
    End If
                    
    If rv0 = 1 Then
        rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
        If rv0 <> 1 Then
            rv0 = 2
            Tester.Print "SD bus width Fail"
        End If
    End If
                    
    ClosePipe
                      
    Tester.Print "rv0="; rv0
                     
    If rv0 = 1 Then
        Do
            Call MsecDelay(0.2)
            CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
            If CardResult <> 0 Then
                MsgBox "Read light On fail"
                End
            End If
    
            'If CAndValue(LightON, &H2) = 2 Then
            If (LightOn <> 252) Then
                If RT_Counter >= 4 Then
                    Tester.Print "LightON="; LightOn
                    UsbSpeedTestResult = GPO_FAIL
                End If
                
                rv0 = 3
            Else
                rv0 = 1
            End If
            
            RT_Counter = RT_Counter + 1
            
        Loop While (rv0 <> 1) And (RT_Counter < 5)
    End If
                    
    If rv0 = 1 Then
        CardResult = DO_WritePort(card, Channel_P1A, &H7D)
        Call MsecDelay(0.04)
        CardResult = DO_WritePort(card, Channel_P1A, &H7E)
        
        rv0 = CBWTest_New(0, 1, ChipString)
          
          
        If rv0 = 1 Then
            rv0 = CBWTest_New_128_Sector_PipeReady(0, rv0)    ' write
        End If
        ClosePipe
    End If
    
    If rv0 = 1 Then
        CardResult = DO_WritePort(card, Channel_P1A, &H7D)
        Call MsecDelay(0.04)
        CardResult = DO_WritePort(card, Channel_P1A, &H7E)
        
        rv0 = CBWTest_New(0, 1, ChipString)
          
          
        If rv0 = 1 Then
            rv0 = CBWTest_New_128_Sector_PipeReady(0, rv0)    ' write
        End If
        ClosePipe
    End If
                
    Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
    Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    '===============================================
    '  CF Card test
    '================================================
               
    rv1 = rv0       '----------- AU6438BS no CF slot
                 
    'Call LabelMenu(1, rv1, rv0)
    'Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    
    '===============================================
    '  SMC Card test  : stop these test for card not enough
    '================================================
                
    rv2 = rv1       '----------- AU6438BS no SMC slot
               
              
    '===============================================
    '  XD Card test
    '================================================
                 
    rv3 = rv2       '----------- AU6438BS no XD slot
                
    
    '===============================================
    '  MS Card test
    '================================================
                     
    rv4 = rv3       '----------- AU6438BS no MS slot
               
    'Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    
    
    '===============================================
    'MS Pro Card test
    '================================================
              
    CardResult = DO_WritePort(card, Channel_P1A, &H5E)
                
    If CardResult <> 0 Then
        MsgBox "Set MSPro Card Detect On Fail"
        End
    End If
                
    Call MsecDelay(0.02)
                
    If rv0 = 1 Then
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
                       
    End If
                
    Call LabelMenu(31, rv5, rv0)
    Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    
'
'    If HV_Flag = False Then
'        HV_Flag = True
'        Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
'        'CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
'        Call MsecDelay(0.3)
'        HV_Result = rv0 * rv5
'        rv0 = 0
'        rv5 = 0
'        GoTo ReTest_LABEL
'    ElseIf HV_Flag = True Then
'        LV_Result = rv0 * rv5
'        Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
'        CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
'        Call MsecDelay(0.1)
'    End If
                
                
AU6371ELResult:
    
    ClosePipe
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
    'WaitDevOFF ("pid_6366")
    SetSiteStatus (SiteUnknow)
    
    If HV_Done_Flag = False Then
        If rv0 <> 1 Then
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
        Call MsecDelay(0.4)
        ReaderExist = 0
        GoTo Routine_Label
    Else
        If rv0 <> 1 Then
            LV_Result = "Bin2"
            Tester.Print "LV Unknow"
        ElseIf rv0 * rv1 * rv2 * rv3 <> 1 Then
            LV_Result = "Fail"
            Tester.Print "LV Fail"
        ElseIf rv0 * rv1 * rv2 * rv3 = 1 Then
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
                   
                
End Sub

Public Sub AU6438BLF20TestSub()
'Tester.Print "AU6433HF: Normal mode "
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
                    LightOn = 0
                    LightOff = 0
                    
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
                       
               
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H1)
                  
                 Call MsecDelay(0.2)     'power on time
                 
               
                '===============================================
                '  SD Card test
                '===============================================
                  
             
                 '============================================
                 'NO card test
                 '============================================
  
                     ' set SD card detect down
                      CardResult = DO_WritePort(card, Channel_P1A, &HDE)    'power-on
                      
                     If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                     End If
                     Call MsecDelay(0.1)
                     Call WaitDevOn(ChipString)
                        
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                          
                      ClosePipe
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      
                      'rv0 = CBWTest_New(0, 1, ChipString)
                      'rv0 = CBWTest_New(0, 1, ChipString)
                      'rv0 = CBWTest_New(0, 1, ChipString)
                      
                      If rv0 = 1 Then
                       rv0 = Read_SD_Speed(0, 0, 64, "4Bits")
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
              
                
                
                CardResult = DO_WritePort(card, Channel_P1A, &HD6)  'SD + MS
                Call MsecDelay(0.03)
                CardResult = DO_WritePort(card, Channel_P1A, &HF6)  'SD
                
                OpenPipe
                rv5 = ReInitial(0)
                ClosePipe
                
                If rv0 = 1 Then
                    'CardResult = DO_WritePort(card, Channel_P1A, &HF6)
                    'Call MsecDelay(0.2)
             
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
                  
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                       
                  CardResult = DO_WritePort(card, Channel_P1A, &HFE)   ' Revmove MS
                    Call MsecDelay(0.2)
             
                  
                    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                        
                     
                    If rv0 <> 0 Then
                        If (CAndValue(LightOn, &H1) <> 0) Or (CAndValue(LightOff, &H1) <> 1) Then
                            Tester.Print "LightON="; LightOn
                            Tester.Print "LightOFF="; LightOff
                            UsbSpeedTestResult = GPO_FAIL
                            rv0 = 3
                        End If
                    End If
                    
                  
                  If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                
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
Public Sub AU6438CLF20TestSub()

Tester.Print "AU6438CL: Normal Mode Begin Test ..."

'==================================================================
'
'  this code come from AU6371BLF20TestSub
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
    LightOn = 0
    LightOff = 0
                    
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
    '    POWER On
    '=========================================
        ChipString = "vid_058f"
        CardResult = DO_WritePort(card, Channel_P1A, &H7E)
        Call MsecDelay(0.2)     'power on time
        
        
    '===============================================
    '  SD Card test
    '===============================================
    
    ' set SD card detect down
        CardResult = DO_WritePort(card, Channel_P1A, &HDE)    'ENA_ON SD_ON
                       
        If CardResult <> 0 Then
            MsgBox "Set SD Card Detect Down Fail"
            End
        End If
    
        Call MsecDelay(0.1)
        rv0 = WaitDevOn(ChipString)
                        
        CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
        If CardResult <> 0 Then
            MsgBox "Read light On fail"
            End
        End If
        
        ClosePipe
        
        If rv0 = 1 Then
            rv0 = CBWTest_New(0, rv0, ChipString)
                      
            If rv0 = 1 Then
                rv0 = Read_SD_Speed(0, 0, 64, "4Bits")
                If rv0 <> 1 Then
                    rv0 = 2
                    Tester.Print "SD bus width Fail"
                End If
            End If
        End If
        
        ClosePipe
        
        If rv0 = 1 Then
            rv0 = CBWTest_New_128_Sector_AU6377(0, rv0)  ' write
                        
            If rv0 = 1 Then
                Tester.Print "128 Sector Test PASS"
            Else
                Tester.Print "128 Sector Test FAIL"
            End If
        End If
        
        ClosePipe
        
        Call LabelMenu(0, rv0, 1)   ' no card test fail
        Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    
    '===============================================
    '  CF Card test
    '================================================
               
        rv1 = rv0  '----------- AU6438 no CF slot
        
        'Call LabelMenu(1, rv1, rv0)
        'Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    '===============================================
    '  SMC Card test  : stop these test for card not enough
    '================================================

        rv2 = rv1   '----------- AU6438 no SMC slot
               
              
    '===============================================
    '  XD Card test
    '================================================
        
        rv3 = rv2   '----------- AU6438 no XD slot
                
    '===============================================
    '  MS Card test
    '================================================
                   
        rv4 = rv3  '----------- AU6438 no MSPro slot
               
        'Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    
    '===============================================
    '  MS Pro Card test
    '================================================
                
        CardResult = DO_WritePort(card, Channel_P1A, &HD6)  'SD + MS
        Call MsecDelay(0.03)
        CardResult = DO_WritePort(card, Channel_P1A, &HF6)  'SD
                
        OpenPipe
        rv5 = ReInitial(0)
        ClosePipe
                
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
                  
             
        CardResult = DO_WritePort(card, Channel_P1A, &HFE)   ' Revmove MS
        Call MsecDelay(0.4)
                  
        CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                     
        If rv0 <> 0 Then
            If (CAndValue(LightOn, &H1) <> 0) Or (CAndValue(LightOff, &H1) <> 1) Then
                Tester.Print "LightON="; LightOn
                Tester.Print "LightOFF="; LightOff
                UsbSpeedTestResult = GPO_FAIL
                rv0 = 3
            End If
        End If
                  
                
AU6371ELResult:
        
        CardResult = DO_WritePort(card, Channel_P1A, &H81)
        
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

Public Sub AU6438CFF20TestSub()
'Tester.Print "AU6433HF: Normal mode "
'==================================================================
'
'  this code come from AU6438BSF20TestSub
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
                    ChipString = "vid_058f"
                       
                    CardResult = DO_WritePort(card, Channel_P1A, &HBF) 'Reset Flip-Flop
                    
                    Call MsecDelay(0.1)
                    
                    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                    
                    If (LightOn And &H2) <> &H2 Then
                        MsgBox ("Reset Flip-Flop IC Fail")
                        End
                    End If
                    
                 'CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                
                '===============================================
                '  SD Card test
                '
             
                     ' set SD card detect down
                    CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                    Call MsecDelay(0.1)     'power on time
             
                      
                    If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                    End If
                    
                    'Call MsecDelay(0.4)
                     
                           
                    ClosePipe
                      
                      
                    rv0 = CBWTest_New(0, 1, ChipString)
                      
                    If rv0 = 1 Then
                       rv0 = Read_SD_Speed(0, 0, 64, "4Bits")   'for AU6438CFF 4Bit SD
                       
                       If rv0 <> 1 Then
                          rv0 = 2
                          Tester.Print "SD bus width Fail"
                       End If
                    End If
                       
                    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                       
                    ClosePipe
                      
                    Tester.Print "rv0="; rv0
                     
                    If rv0 <> 0 Then
                        If (LightOn And &H2) <> 0 Then      'just compare Bit2
                            Tester.Print "LightON="; LightOn
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
              
                rv5 = 1 'AU6438JCF has no MSpro slot
                
                  
                
AU6371ELResult:

                CardResult = DO_WritePort(card, Channel_P1A, &H80)   ' Close power
                Call MsecDelay(0.2)
                
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

Public Sub AU6438CFS10TestSub()

'==================================================================
'  this code come from AU6438CFF20TestSub
'===================================================================
'2012/10/24: purpose to soting RFTECH RMA sample
'2014/3/25: Skip this flow


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
                
'=========================================
'    POWER on
'=========================================
ChipString = "vid_058f"

Call PowerSet2(1, "4.4", "0.3", 1, "4.4", "0.3", 1)
Call MsecDelay(0.2)

CardResult = DO_WritePort(card, Channel_P1A, &HBF) 'Reset Flip-Flop
Call MsecDelay(0.1)

CardResult = DO_ReadPort(card, Channel_P1B, LightOn)

If (LightOn And &H2) <> &H2 Then
    MsgBox ("Reset Flip-Flop IC Fail")
    End
End If
                    
'CardResult = DO_WritePort(card, Channel_P1A, &H7E)

'===============================================
'  SD Card test
'
             
' set SD card detect down
CardResult = DO_WritePort(card, Channel_P1A, &H7E)
Call MsecDelay(0.2)     'power on time
rv0 = WaitDevOn(ChipString)

If rv0 = 1 Then
    ReaderExist = 1
Else
    ReaderExist = 0
End If

 
If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If
Call MsecDelay(0.2)
                     
ClosePipe
rv0 = CBWTest_New(0, 1, ChipString)


  
If rv0 = 1 Then
    
    If rv0 = 1 Then
        rv0 = CBWTest_New_128_Sector_PipeReady(0, rv0)    ' write
    End If
    
    If rv0 = 1 Then
        rv0 = Read_SD_Speed(0, 0, 64, "4Bits")   'for AU6438CFF 4Bit SD
    End If
    
    If rv0 <> 1 Then
        rv0 = 2
        Tester.Print "SD bus width Fail"
    End If
End If
   
CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
ClosePipe
  
Tester.Print "rv0="; rv0
 
If rv0 <> 0 Then
    If (LightOn And &H2) <> 0 Then      'just compare Bit2
        Tester.Print "LightON="; LightOn
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

 rv5 = 1 'AU6438JCF has no MSpro slot
                
                
AU6371ELResult:

    CardResult = DO_WritePort(card, Channel_P1A, &H80)   ' Close power
    Call MsecDelay(0.2)
                      
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
Public Sub AU6438CFS21TestSub()

'==================================================================
'  this code come from AU6438CFF20TestSub
'===================================================================
'2012/10/24: purpose to soting RFTECH RMA sample

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
                
'=========================================
'    POWER on
'=========================================
ChipString = "vid_058f"

If (AU6438SortingFlag) Then
    Call PowerSet2(1, "4.4", "0.3", 1, "4.4", "0.3", 1)
    Call MsecDelay(0.2)
End If

CardResult = DO_WritePort(card, Channel_P1A, &HBF) 'Reset Flip-Flop
Call MsecDelay(0.1)

CardResult = DO_ReadPort(card, Channel_P1B, LightOn)

'If (LightOn And &H2) <> &H2 Then
'    MsgBox ("Reset Flip-Flop IC Fail")
'    End
'End If
                    
'CardResult = DO_WritePort(card, Channel_P1A, &H7E)

'===============================================
'  SD Card test
'
             
' set SD card detect down
CardResult = DO_WritePort(card, Channel_P1A, &H7E)
Call MsecDelay(0.2)     'power on time
rv0 = WaitDevOn(ChipString)

If rv0 = 1 Then
    ReaderExist = 1
Else
    ReaderExist = 0
End If

 
If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If
Call MsecDelay(0.2)
                     
ClosePipe
rv0 = CBWTest_New(0, 1, ChipString)


  
If rv0 = 1 Then
    
    If rv0 = 1 Then
        rv0 = CBWTest_New_128_Sector_PipeReady(0, rv0)    ' write
    End If
    
    If rv0 = 1 Then
        rv0 = Read_SD_Speed(0, 0, 64, "4Bits")   'for AU6438CFF 4Bit SD
    End If
    
    If rv0 <> 1 Then
        rv0 = 2
        Tester.Print "SD bus width Fail"
    End If
End If
   
CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
ClosePipe
  
Tester.Print "rv0="; rv0
 
If rv0 <> 0 Then
    If (LightOn And &H2) <> 0 Then      'just compare Bit2
        Tester.Print "LightON="; LightOn
        UsbSpeedTestResult = GPO_FAIL
        rv0 = 3
    End If
End If

If rv0 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &H7D)
    Call MsecDelay(0.1)
    CardResult = DO_WritePort(card, Channel_P1A, &H7E)
    rv0 = CBWTest_New(0, 1, ChipString)
    
    If rv0 = 1 Then
        rv0 = CBWTest_New_128_Sector_PipeReady(0, rv0)    ' write
    End If
    ClosePipe
End If

If rv0 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &H7D)
    Call MsecDelay(0.1)
    CardResult = DO_WritePort(card, Channel_P1A, &H7E)
    rv0 = CBWTest_New(0, 1, ChipString)
    
    If rv0 = 1 Then
        rv0 = CBWTest_New_128_Sector_PipeReady(0, rv0)    ' write
    End If
    ClosePipe
End If

    
Call LabelMenu(0, rv0, 1)   ' no card test fail
 
Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
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

 rv5 = 1 'AU6438JCF has no MSpro slot
                
                
AU6371ELResult:

    CardResult = DO_WritePort(card, Channel_P1A, &H80)   ' Close power
    Call MsecDelay(0.2)
                      
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
Public Sub AU6438CFF00TestSub()
'Tester.Print "AU6433HF: Normal mode "
'==================================================================
'
'  this code come from AU6438BSF20TestSub
'  2011/9/13 HV + LV test
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
    Dim HV_Result As Byte
    Dim LV_Result As Byte
    OldChipName = ""
                 
    ' initial condition
                
    AU6371EL_SD = 1
    AU6371EL_CF = 2
    AU6371EL_XD = 8
    AU6371EL_MS = 32
    AU6371EL_MSP = 64
    AU6371EL_BootTime = 0.6
    
    HV_Flag = False
    HV_Result = 0
    LV_Result = 0
                     
    If PCI7248InitFinish = 0 Then
        PCI7248Exist
    End If
               
    ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
    '  CardResult = DO_WritePort(card, Channel_P1B, &H0)

ReTest_LABEL:


    LBA = LBA + 1
                         
    If HV_Flag = False Then
        Call PowerSet2(1, "5.3", "0.3", 1, "5.3", "0.3", 1)
        Tester.Print "Start HV(5.3) Test ..."
        Call MsecDelay(0.2)
    Else
        Tester.Print vbCrLf & vbCrLf & "Start LV(4.7) Test ..."
        Call PowerSet2(1, "4.7", "0.3", 1, "4.7", "0.3", 1)
        Call MsecDelay(0.2)
    End If
                         
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
                       
    CardResult = DO_WritePort(card, Channel_P1A, &HBF) 'Reset Flip-Flop
                    
    Call MsecDelay(0.1)
                    
    'CardResult = DO_ReadPort(card, Channel_P1B, LightON)
                    
    'If (LightON And &H2) <> &H2 Then
    '    MsgBox ("Reset Flip-Flop IC Fail")
    '    End
    'End If
                    
    'CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                
    '===============================================
    '  SD Card test
    '===============================================
             
    ' set SD card detect down
    CardResult = DO_WritePort(card, Channel_P1A, &H7E)
    Call MsecDelay(0.1)     'power on time
                      
    If CardResult <> 0 Then
        MsgBox "Set SD Card Detect Down Fail"
        End
    End If
                    
    rv0 = WaitDevOn(ChipString)
    Call MsecDelay(0.1)
    
    ClosePipe
                      
    If rv0 = 1 Then
        rv0 = CBWTest_New(0, 1, ChipString)
                      
        If rv0 = 1 Then
            rv0 = Read_SD_Speed(0, 0, 64, "4Bits")   'for AU6438CFF 4Bit SD
            If rv0 <> 1 Then
                rv0 = 2
                Tester.Print "SD bus width Fail"
            End If
        End If
    End If
    
    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                       
    ClosePipe
                      
    Tester.Print "rv0="; rv0
                     
    If rv0 = 1 Then
        If (LightOn And &H2) <> 0 Then      'just compare Bit2
            Tester.Print "LightON="; LightOn
            UsbSpeedTestResult = GPO_FAIL
            rv0 = 3
        End If
    End If
                    
    Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
    Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    '===============================================
    '  CF Card test
    '================================================
               
    rv1 = rv0  '----------- AU6438CF has no CF slot
                 
    'Call LabelMenu(1, rv1, rv0)
    'Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    
    '===============================================
    '  SMC Card test  : stop these test for card not enough
    '================================================
              
    'AU6438 has no SMC slot
               
    rv2 = rv1   ' to complete the SMC asbolish
               
              
    '===============================================
    '  XD Card test
    '================================================
    
    rv3 = rv2
                
    '===============================================
    '  MS Card test
    '================================================
                
    rv4 = rv3  'AU6344 has no MS slot pin
               
    'Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    '===============================================
    '  MS Pro Card test
    '================================================
              
    rv5 = rv4 'AU6438JCF has no MSpro slot
                  
                  
    '=================================================
    '   NO Card test
    '=================================================
                
    CardResult = DO_WritePort(card, Channel_P1A, &H7F)   ' Close power
    Call MsecDelay(0.2)
             
    If GetDeviceName(ChipString) = "" Then
        rv0 = 0
    End If
                 
    
    If HV_Flag = False Then
        HV_Flag = True
        Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
        CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
        Call MsecDelay(0.3)
        HV_Result = rv0
        ReaderExist = 0
        GoTo ReTest_LABEL
    ElseIf HV_Flag = True Then
        LV_Result = rv0
        Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
        CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
        Call MsecDelay(0.05)
    End If
                
                
AU6371ELResult:
                
    If (HV_Result = 0) And (LV_Result = 0) Then
        UnknowDeviceFail = UnknowDeviceFail + 1
        TestResult = "Bin2"
    ElseIf (HV_Result <> 1) And (LV_Result = 1) Then
        TestResult = "Bin3"
    ElseIf (HV_Result = 1) And (LV_Result <> 1) Then
        TestResult = "Bin4"
    ElseIf (HV_Result <> 1) And (LV_Result <> 1) Then
        TestResult = "Bin5"
    ElseIf (HV_Result = 1) And (LV_Result = 1) Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
                
                
End Sub

Public Sub AU6438CFF01TestSub()
'Tester.Print "AU6433HF: Normal mode "
'==================================================================
'
'  this code come from AU6438BSF20TestSub
'  2014/4/15 HV + LV add sd insert/remove test
'
'===================================================================


    Dim ChipString As String
    Dim AU6371EL_SD As Byte
    Dim AU6371EL_CF As Byte
    Dim AU6371EL_XD As Byte
    Dim AU6371EL_MS As Byte
    Dim AU6371EL_MSP  As Byte
    Dim AU6371EL_BootTime As Single
    Dim HV_Done_Flag As Boolean
    Dim HV_Result As String
    Dim LV_Result As String
    OldChipName = ""
                 
    ' initial condition
                
    AU6371EL_SD = 1
    AU6371EL_CF = 2
    AU6371EL_XD = 8
    AU6371EL_MS = 32
    AU6371EL_MSP = 64
    AU6371EL_BootTime = 0.6
    
    HV_Done_Flag = False
    HV_Result = ""
    LV_Result = ""
                     
    If PCI7248InitFinish_Sync = 0 Then
        PCI7248Exist_P1C_Sync
    End If
               
    ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
    '  CardResult = DO_WritePort(card, Channel_P1B, &H0)

Routine_Label:


    LBA = LBA + 1
                         
    If HV_Done_Flag = False Then
        Call PowerSet2(1, "5.25", "0.3", 1, "5.25", "0.3", 1)
        Tester.Print "Start HV(5.25) Test ..."
        SetSiteStatus (RunHV)
        'Call MsecDelay(0.2)
    Else
        Tester.Print vbCrLf & vbCrLf & "Start LV(4.4) Test ..."
        Call PowerSet2(1, "4.4", "0.3", 1, "4.4", "0.3", 1)
        SetSiteStatus (RunLV)
        'Call MsecDelay(0.2)
    End If
                         
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
    ChipString = "pid_6366"
                       
    CardResult = DO_WritePort(card, Channel_P1A, &HBF) 'Reset Flip-Flop
                    
    Call MsecDelay(0.1)
                    
    'CardResult = DO_ReadPort(card, Channel_P1B, LightON)
                    
    'If (LightON And &H2) <> &H2 Then
    '    MsgBox ("Reset Flip-Flop IC Fail")
    '    End
    'End If
                    
    'CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                
    '===============================================
    '  SD Card test
    '===============================================
             
    ' set SD card detect down
    CardResult = DO_WritePort(card, Channel_P1A, &H7E)
    Call MsecDelay(0.1)     'power on time
                      
    If CardResult <> 0 Then
        MsgBox "Set SD Card Detect Down Fail"
        End
    End If
                    
    rv0 = WaitDevOn(ChipString)
    Call MsecDelay(0.1)
    
    ClosePipe
                      
    If rv0 = 1 Then
        rv0 = CBWTest_New(0, 1, ChipString)
                      
        If rv0 = 1 Then
            rv0 = Read_SD_Speed(0, 0, 64, "4Bits")   'for AU6438CFF 4Bit SD
            If rv0 <> 1 Then
                rv0 = 2
                Tester.Print "SD bus width Fail"
            End If
        End If
    End If
    
    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                       
    ClosePipe
                      
    Tester.Print "rv0="; rv0
                     
    If rv0 = 1 Then
        If (LightOn And &H2) <> 0 Then      'just compare Bit2
            Tester.Print "LightON="; LightOn
            UsbSpeedTestResult = GPO_FAIL
            rv0 = 3
        End If
    End If
    
    If rv0 = 1 Then
        CardResult = DO_WritePort(card, Channel_P1A, &H7D)
        Call MsecDelay(0.1)
        CardResult = DO_WritePort(card, Channel_P1A, &H7E)
        rv0 = CBWTest_New(0, 1, ChipString)
        
        If rv0 = 1 Then
            rv0 = CBWTest_New_128_Sector_PipeReady(0, rv0)    ' write
        End If
        ClosePipe
    End If

    If rv0 = 1 Then
        CardResult = DO_WritePort(card, Channel_P1A, &H7D)
        Call MsecDelay(0.1)
        CardResult = DO_WritePort(card, Channel_P1A, &H7E)
        rv0 = CBWTest_New(0, 1, ChipString)
        
        If rv0 = 1 Then
            rv0 = CBWTest_New_128_Sector_PipeReady(0, rv0)    ' write
        End If
        ClosePipe
    End If
                    
    Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
    Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    '===============================================
    '  CF Card test
    '================================================
               
    rv1 = rv0  '----------- AU6438CF has no CF slot
                 
    'Call LabelMenu(1, rv1, rv0)
    'Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    
    '===============================================
    '  SMC Card test  : stop these test for card not enough
    '================================================
              
    'AU6438 has no SMC slot
               
    rv2 = rv1   ' to complete the SMC asbolish
               
              
    '===============================================
    '  XD Card test
    '================================================
    
    rv3 = rv2
                
    '===============================================
    '  MS Card test
    '================================================
                
    rv4 = rv3  'AU6344 has no MS slot pin
               
    'Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    '===============================================
    '  MS Pro Card test
    '================================================
              
    rv5 = rv4 'AU6438JCF has no MSpro slot
                  
                  
    '=================================================
    '   NO Card test
    '=================================================
                
    CardResult = DO_WritePort(card, Channel_P1A, &H7F)   ' Close power
    Call MsecDelay(0.2)
             
    If GetDeviceName(ChipString) = "" Then
        rv0 = 0
    End If
               
                
AU6371ELResult:

    ClosePipe
    CardResult = DO_WritePort(card, Channel_P1A, &H80)
    If Not HV_Done_Flag Then
        SetSiteStatus (HVDone)
        Call WaitAnotherSiteDone(HVDone, 3#)
    Else
        SetSiteStatus (LVDone)
        Call WaitAnotherSiteDone(LVDone, 3#)
    End If
    Call PowerSet2(0, "0.0", "0.5", 1, "0.0", "0.5", 1)
    WaitDevOFF (ChipString)
    Call MsecDelay(0.1)
    SetSiteStatus (SiteUnknow)

                
    If HV_Done_Flag = False Then
        If rv0 <> 1 Then
            HV_Result = "Bin2"
            Tester.Print "HV Unknow"
        ElseIf rv0 * rv1 * rv2 * rv3 <> 1 Then
            HV_Result = "Fail"
            Tester.Print "HV Fail"
        ElseIf rv0 * rv1 * rv2 * rv3 = 1 Then
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
        ElseIf rv0 * rv1 * rv2 * rv3 <> 1 Then
            LV_Result = "Fail"
            Tester.Print "LV Fail"
        ElseIf rv0 * rv1 * rv2 * rv3 = 1 Then
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
                
End Sub


Public Sub AU6438EFS20TestSub()

'20140325: Revise from AU6438EFF20TestSub
'          This flow add 36 ohm on DP, DM & using external 4.4V


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
LightOn = 0
LightOff = 0
     
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


ChipString = "vid_058f"

If (AU6438SortingFlag) Then
    Call PowerSet2(1, "4.4", "0.3", 1, "4.4", "0.3", 1)
End If

' PortA        PortB
'-----------------------
'E M    S            G
'N S    D            P
'A p                 O
'  r                 N
'  o                 7
'||||||||     ||||||||
'||||||||     ||||||||
'87654321     87654321
                       
                    
'CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
'Call MsecDelay(0.01)

CardResult = DO_WritePort(card, Channel_P1A, &HDE)  'ENA=> "OFF", MS-SD=> "ON"

Call MsecDelay(0.1)
                    
'=========================================
'    POWER on
'=========================================
    
'===============================================
'  SD Card test
'
             
 ' set SD card detect down
CardResult = DO_WritePort(card, Channel_P1A, &H7E)
WaitDevOn (ChipString)      'power on time

CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
Call MsecDelay(0.01)
 
If (LightOn <> 254) Then
    Tester.Print "LightON="; LightOn
    UsbSpeedTestResult = GPO_FAIL
    rv0 = 3
    Call LabelMenu(0, rv0, 1)
    GoTo AU6438EFResult
End If


If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If

rv0 = CBWTest_New(0, 1, ChipString)

If rv0 = 1 Then
    rv0 = Read_SD_Speed(0, 0, 64, "8Bits")   'for AU6438CFF 8Bit 48MHz SD
   
    ClosePipe
    
    If rv0 <> 1 Then
        rv0 = 2
        Call LabelMenu(0, rv0, 1)
        Tester.Print "SD bus width Fail"
        GoTo AU6438EFResult
    End If
    
    If rv0 = 1 Then
        rv0 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
    
        If rv0 = 1 Then
            Tester.Print "128 Sector Test PASS"
        Else
            Call LabelMenu(0, rv0, 1)
            Tester.Print "128 Sector Test FAIL"
            GoTo AU6438EFResult
        End If
    End If
    ClosePipe
End If

If rv0 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &H7F)
    Call MsecDelay(0.1)
    CardResult = DO_WritePort(card, Channel_P1A, &H7E)
    rv0 = CBWTest_New(0, 1, ChipString)
    If rv0 = 1 Then
        rv0 = CBWTest_New_128_Sector_PipeReady(0, rv0)    ' write
    End If
    ClosePipe
End If


ClosePipe
Call LabelMenu(0, rv0, 1)   ' no card test fail

Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

'=================================================
'   NO Card test
'=================================================

CardResult = DO_WritePort(card, Channel_P1A, &H7F)   ' Close power
Call MsecDelay(0.2)

If GetDeviceName(ChipString) = "" Then
    rv0 = 0
    Call LabelMenu(0, rv0, 1)
    GoTo AU6438EFResult
End If
                
'===============================================
'  CF Card test
'================================================

rv1 = rv0  '----------- AU6438EF has not CF slot
 
 
'  Call LabelMenu(1, rv1, rv0)

'   Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
'===============================================
'  SMC Card test  : stop these test for card not enough
'================================================

'AU6438EF has no SMC slot

rv2 = 1   ' to complete the SMC asbolish


'===============================================
'  XD Card test
'================================================
rv3 = 1

'===============================================
'  MS Card test
'================================================
   
     
rv4 = 1  'AU6438EF has no MS slot pin

'    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
'===============================================
'  MS Pro Card test
'================================================

OpenPipe
rv5 = ReInitial(0)
ClosePipe

CardResult = DO_WritePort(card, Channel_P1A, &H5F)  'open MS
Call MsecDelay(0.1)


rv5 = CBWTest_New(0, rv4, ChipString)
 
If rv5 = 1 Then
    rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
    
    If rv5 <> 1 Then
        rv5 = 2
        Tester.Print "MS bus width Fail"
    End If
End If

Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
ClosePipe
  
Call LabelMenu(31, rv5, rv4)
                
AU6438EFResult:
                      
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)  'Close All
    WaitDevOFF (ChipString)
    
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

Public Sub AU6438MFS20TestSub()

'20140325: Revise from AU6438EFF20TestSub
'          This flow add 36 ohm on DP, DM & using external 4.4V


Dim ChipString As String

Dim AU6371EL_SD As Byte
Dim AU6371EL_CF As Byte
Dim AU6371EL_XD As Byte
Dim AU6371EL_MS As Byte
Dim AU6371EL_MSP  As Byte
Dim AU6371EL_BootTime As Single
Dim k As Integer
OldChipName = ""

' initial condition
 
AU6371EL_SD = 1
AU6371EL_CF = 2
AU6371EL_XD = 8
AU6371EL_MS = 32
AU6371EL_MSP = 64
AU6371EL_BootTime = 0.6
LightOn = 0
LightOff = 0
     
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


ChipString = "vid_058f"

If (AU6438SortingFlag) Then
    Call PowerSet2(1, "4.4", "0.3", 1, "4.4", "0.3", 1)
End If

' PortA        PortB
'-----------------------
'E M    S            G
'N S    D            P
'A p                 O
'  r                 N
'  o                 7
'||||||||     ||||||||
'||||||||     ||||||||
'87654321     87654321
                       
                    
'CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
'Call MsecDelay(0.01)

CardResult = DO_WritePort(card, Channel_P1A, &HDE)  'ENA=> "OFF", MS-SD=> "ON"

Call MsecDelay(0.1)
                    
'=========================================
'    POWER on
'=========================================
    
'===============================================
'  SD Card test
'
             
 ' set SD card detect down
CardResult = DO_WritePort(card, Channel_P1A, &H7E)
WaitDevOn (ChipString)      'power on time

For k = 1 To 10
    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
    Call MsecDelay(0.01)
    If (LightOn And &H1) = &H0 Then
        Exit For
    End If
Next

If ((LightOn And &H1) <> &H0) Then
    Tester.Print "LightON="; LightOn
    UsbSpeedTestResult = GPO_FAIL
    rv0 = 3
    Call LabelMenu(0, rv0, 1)
    GoTo AU6438EFResult
End If


If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If

rv0 = CBWTest_New(0, 1, ChipString)

If rv0 = 1 Then
    rv0 = Read_SD_Speed(0, 0, 64, "8Bits")   'for AU6438CFF 8Bit 48MHz SD
   
    ClosePipe
    
    If rv0 <> 1 Then
        rv0 = 2
        Call LabelMenu(0, rv0, 1)
        Tester.Print "SD bus width Fail"
        GoTo AU6438EFResult
    End If
    
    If rv0 = 1 Then
        rv0 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
    
        If rv0 = 1 Then
            Tester.Print "128 Sector Test PASS"
        Else
            Call LabelMenu(0, rv0, 1)
            Tester.Print "128 Sector Test FAIL"
            GoTo AU6438EFResult
        End If
    End If
    ClosePipe
End If

If rv0 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &H7F)
    Call MsecDelay(0.1)
    CardResult = DO_WritePort(card, Channel_P1A, &H7E)
    rv0 = CBWTest_New(0, 1, ChipString)
    If rv0 = 1 Then
        rv0 = CBWTest_New_128_Sector_PipeReady(0, rv0)    ' write
    End If
    ClosePipe
End If


ClosePipe
Call LabelMenu(0, rv0, 1)   ' no card test fail

Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

'=================================================
'   NO Card test
'=================================================

CardResult = DO_WritePort(card, Channel_P1A, &H7F)   ' Close power
Call MsecDelay(0.2)

If GetDeviceName(ChipString) = "" Then
    rv0 = 0
    Call LabelMenu(0, rv0, 1)
    GoTo AU6438EFResult
End If
                
'===============================================
'  CF Card test
'================================================

rv1 = rv0  '----------- AU6438EF has not CF slot
 
 
'  Call LabelMenu(1, rv1, rv0)

'   Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
'===============================================
'  SMC Card test  : stop these test for card not enough
'================================================

'AU6438EF has no SMC slot

rv2 = 1   ' to complete the SMC asbolish


'===============================================
'  XD Card test
'================================================
rv3 = 1

'===============================================
'  MS Card test
'================================================
   
     
rv4 = 1  'AU6438EF has no MS slot pin

'    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
'===============================================
'  MS Pro Card test
'================================================
                
'OpenPipe
'rv5 = ReInitial(0)
'ClosePipe
'
'CardResult = DO_WritePort(card, Channel_P1A, &H5F)  'open MS
'Call MsecDelay(0.1)
'
'
'rv5 = CBWTest_New(0, rv4, ChipString)
'
'If rv5 = 1 Then
'    rv5 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
'
'    If rv5 <> 1 Then
'        rv5 = 2
'        Tester.Print "MS bus width Fail"
'    End If
'End If
'
'Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
'ClosePipe
'
'Call LabelMenu(31, rv5, rv4)
                
AU6438EFResult:
                      
    CardResult = DO_WritePort(card, Channel_P1A, &H80)  'Close All
    WaitDevOFF (ChipString)
      
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
    ElseIf rv0 = PASS Then
         TestResult = "PASS"
    Else
        TestResult = "Bin2"
      
    End If
    
End Sub

Public Sub AU6438GFF20TestSub()

'==================================================================
'
'  this code come from AU6438EFF20TestSub
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
                    LightOn = 0
                    LightOff = 0
                    
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
                
                
                    ChipString = "vid_058f"
                       
                    ' PortA        PortB
                    '-----------------------
                    'E      S            G
                    'N      D            P
                    'A                   O
                    '                    N
                    '                    7
                    '||||||||     ||||||||
                    '||||||||     ||||||||
                    '87654321     87654321
                       
                    
                    'CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
                    'Call MsecDelay(0.01)
                    
                    CardResult = DO_WritePort(card, Channel_P1A, &HDE)  'ENA=> "OFF", SD=> "ON"
                    
                    Call MsecDelay(0.1)
                    
                '=========================================
                '    POWER on
                '=========================================
                    
                '===============================================
                '  SD Card test
                '
             
                     ' set SD card detect down
                    CardResult = DO_WritePort(card, Channel_P1A, &H7E)
                    Call MsecDelay(0.3)       'power on time
                    
                    rv0 = WaitDevOn("vid")
                    
                    If rv0 = 1 Then
                        CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                        Call MsecDelay(0.01)
                     
                        If (LightOn <> 254) Then
                            Tester.Print "LightON="; LightOn
                            UsbSpeedTestResult = GPO_FAIL
                            rv0 = 3
                            Call LabelMenu(0, rv0, 1)
                            GoTo AU6438GFResult
                        End If
                    
                    End If
                    
                    If CardResult <> 0 Then
                        MsgBox "Set SD Card Detect Down Fail"
                        End
                    End If
                    
                    If rv0 = 1 Then
                        rv0 = CBWTest_New(0, 1, ChipString)
                    End If
                    
                    
                    If rv0 = 1 Then
                       rv0 = Read_SD_Speed(0, 0, 64, "8Bits")   'for AU6438CFF 8Bit 48MHz SD
                       
                        ClosePipe
                        
                        If rv0 <> 1 Then
                            rv0 = 2
                            Call LabelMenu(0, rv0, 1)
                            Tester.Print "SD bus width Fail"
                            GoTo AU6438GFResult
                        End If
                        
                        If rv0 = 1 Then
                            rv0 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                        
                            If rv0 = 1 Then
                                Tester.Print "128 Sector Test PASS"
                            Else
                                Call LabelMenu(0, rv0, 1)
                                Tester.Print "128 Sector Test FAIL"
                                GoTo AU6438GFResult
                            End If
                        End If
                    End If
                       
                    ClosePipe
                    Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                    Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '=================================================
                '   NO Card test
                '=================================================
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H7F)   ' Close power
                    Call MsecDelay(0.1)
             
                    If GetDeviceName(ChipString) <> "" Then
                        rv0 = 0
                        Call LabelMenu(0, rv0, 1)
                        GoTo AU6438GFResult
                    End If
                
                '===============================================
                '  CF Card test
                '================================================
               
                rv1 = rv0  '----------- AU6438GF has not CF slot
                 
                 
               '  Call LabelMenu(1, rv1, rv0)
            
                   '   Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6438GF has no SMC slot
               
                rv2 = rv1   ' to complete the SMC asbolish
               
              
                '===============================================
                '  XD Card test
                '================================================
                rv3 = rv2
                
                '===============================================
                '  MS Card test
                '================================================
                   
                     
                rv4 = rv3  'AU6438GF has no MS slot pin
               
                 '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  MS Pro Card test
                '================================================
                
                rv5 = rv4   'A6438GF has no MSpro slot
                
                
AU6438GFResult:
                      
                      CardResult = DO_WritePort(card, Channel_P1A, &HFF)  'Close All
                      
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

Public Sub AU6438IFF20TestSub()
      
Tester.Print "AU6438IF Test ..."

'==================================================================
'
'  this code come from AU6433EFF23 TestSub
'
'===================================================================

Dim ChipString As String
Dim AU6371EL_SD As Byte
Dim AU6371EL_CF As Byte
Dim AU6371EL_XD As Byte
Dim AU6371EL_MS As Byte
Dim AU6371EL_MSP  As Byte
Dim AU6371EL_BootTime As Single
Dim ReadLEDCounter As Integer

OldChipName = ""
                
'initial condition
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
'   POWER on
'   SD Card test
'=========================================
ChipString = "vid"

'CardResult = DO_WritePort(card, Channel_P1A, &H7F)
'Call MsecDelay(0.2)

CardResult = DO_WritePort(card, Channel_P1A, &H7E)

If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If

Call MsecDelay(0.2)
rv0 = WaitDevOn(ChipString)
Call MsecDelay(0.1)

ReadLEDCounter = 0

Do
    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
    If LightOn <> 191 Then
        Call MsecDelay(0.05)
        ReadLEDCounter = ReadLEDCounter + 1
    End If
Loop Until (ReadLEDCounter > 3) Or (LightOn = 191)

If CardResult <> 0 Then
    MsgBox "Read light On fail"
    End
End If

ClosePipe

rv0 = CBWTest_New(0, 1, ChipString)
                      
If rv0 = 1 Then
    rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
    ClosePipe
    If rv0 <> 1 Then
        rv0 = 2
        Tester.Print "SD bus width Fail"
    End If
End If

If rv0 = 1 Then
    rv0 = CBWTest_New_128_Sector_AU6377(0, 1)
    ClosePipe
End If

Tester.Print "rv0="; rv0

CardResult = DO_WritePort(card, Channel_P1A, &H7F)
Call MsecDelay(0.2)

If GetDeviceName(ChipString) <> "" Then
    rv0 = 0
    Tester.Print "NBMD Test Fail"
    GoTo AU6371ELResult
End If
                        
CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
If CardResult <> 0 Then
    MsgBox "Read light off fail"
    End
End If
                     
If rv0 = 1 Then
    If (LightOn <> &HBF) Or (LightOff <> &HFF) Then
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
               
rv1 = rv0       'AU6438GIF not have CF slot
'Call LabelMenu(1, rv1, rv0)
            
'Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                

'===============================================
'  SMC Card test  : stop these test for card not enough
'================================================
              
rv2 = rv1       'AU6438GIF has no SMC slot
'Call LabelMenu(1, rv2, rv1)
                      
                
'===============================================
'  XD Card test
'================================================

rv3 = rv2       'AU6438GIF has no XD slot

'Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                
'===============================================
'  MS Card test
'================================================
                     
rv4 = rv3       'AU6344 has no MS slot pin
               
'Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'===============================================
'  MS Pro Card test
'================================================
                
rv5 = rv4       'AU6344 has no MSpro slot pin
                
'Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                 
                
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

Public Sub AU6438IFS30TestSub()
      
'2011/7/15 add FT6
'2014/03/25: Add ST2 condition (36 ohm on DP,DM & 4.4V)


'==================================================================
'
'  this code come from AU6438IFF20 TestSub
'
'===================================================================

Dim ChipString As String
Dim AU6371EL_SD As Byte
Dim AU6371EL_CF As Byte
Dim AU6371EL_XD As Byte
Dim AU6371EL_MS As Byte
Dim AU6371EL_MSP  As Byte
Dim AU6371EL_BootTime As Single
Dim ReadLEDCounter As Integer

OldChipName = ""
                
'initial condition
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
                

'==========================================================
'    Start Open Shot test
'==========================================================

OS_Result = 0
rv0 = 0

CardResult = DO_WritePort(card, Channel_P1C, &H0)
                 
MsecDelay (0.3)

OpenShortTest_Result_AU6438IFF30

If OS_Result <> 1 Then
    rv0 = 0                 'OS Fail
    GoTo AU6371ELResult
End If

CardResult = DO_WritePort(card, Channel_P1C, &HFF)

Tester.Print "AU6438IF Test ..."


'=========================================
'   POWER on
'   SD Card test
'=========================================
ChipString = "vid_058f"

If (AU6438SortingFlag) Then
    Call PowerSet2(1, "4.4", "0.3", 1, "4.4", "0.3", 1)
End If

CardResult = DO_WritePort(card, Channel_P1A, &H8E)
Call MsecDelay(0.2)

CardResult = DO_WritePort(card, Channel_P1A, &H7E)

If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If

Call MsecDelay(0.2)
rv0 = WaitDevOn(ChipString)
Call MsecDelay(0.1)

ReadLEDCounter = 0

Do
    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
    If LightOn <> 191 Then
        Call MsecDelay(0.05)
        ReadLEDCounter = ReadLEDCounter + 1
    End If
Loop Until (ReadLEDCounter > 3) Or (LightOn = 191)

If CardResult <> 0 Then
    MsgBox "Read light On fail"
    End
End If

ClosePipe
rv0 = CBWTest_New(0, 1, ChipString)
                      
If rv0 = 1 Then
    rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
    ClosePipe
    If rv0 <> 1 Then
        rv0 = 2
        Tester.Print "SD bus width Fail"
    End If
End If

If rv0 = 1 Then
    rv0 = CBWTest_New_128_Sector_AU6377(0, 1)
    ClosePipe
End If

Tester.Print "rv0="; rv0


CardResult = DO_WritePort(card, Channel_P1A, &H7F)
Call MsecDelay(0.2)

If GetDeviceName(ChipString) <> "" Then
    rv0 = 0
    Tester.Print "NBMD Test Fail"
    GoTo AU6371ELResult
End If
                        
CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
If CardResult <> 0 Then
    MsgBox "Read light off fail"
    End
End If
                     
If rv0 = 1 Then
    If (LightOn <> &HBF) Or (LightOff <> &HFF) Then
        Tester.Print "LightON="; LightOn
        Tester.Print "LightOFF="; LightOff
        UsbSpeedTestResult = GPO_FAIL
        rv0 = 3
    End If
End If
                     
Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
If rv0 = 0 Then
    rv0 = 2 'Bin2 Only O/S fail
End If
                
'===============================================
'  CF Card test
'================================================
               
rv1 = rv0       'AU6438GIF not have CF slot
'Call LabelMenu(1, rv1, rv0)
            
'Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                

'===============================================
'  SMC Card test  : stop these test for card not enough
'================================================
              
rv2 = rv1       'AU6438GIF has no SMC slot
'Call LabelMenu(1, rv2, rv1)
                      
                
'===============================================
'  XD Card test
'================================================

rv3 = rv2       'AU6438GIF has no XD slot

'Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                
'===============================================
'  MS Card test
'================================================
                     
rv4 = rv3       'AU6344 has no MS slot pin
               
'Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'===============================================
'  MS Pro Card test
'================================================
                
rv5 = rv4       'AU6344 has no MSpro slot pin
                
'Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                 
                
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


Public Sub AU6438KFS30TestSub()
      
'2012/8/9 for AU6438R64-RKF non support MMC card
'2011/7/15 add FT6
'2014/03/25: Add ST2 condition (36 ohm on DP,DM & 4.4V)



'==================================================================
'
'  this code come from AU6438IFF30 TestSub
'
'===================================================================

Dim ChipString As String
Dim AU6371EL_SD As Byte
Dim AU6371EL_CF As Byte
Dim AU6371EL_XD As Byte
Dim AU6371EL_MS As Byte
Dim AU6371EL_MSP  As Byte
Dim AU6371EL_BootTime As Single
Dim ReadLEDCounter As Integer

OldChipName = ""
                
'initial condition
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
                

'==========================================================
'    Start Open Shot test
'==========================================================

OS_Result = 0
rv0 = 0

CardResult = DO_WritePort(card, Channel_P1C, &H0)
                 
MsecDelay (0.3)

OpenShortTest_Result_AU6438IFF30

If OS_Result <> 1 Then
    rv0 = 0                 'OS Fail
    GoTo AU6371ELResult
End If

CardResult = DO_WritePort(card, Channel_P1C, &HFF)

Tester.Print "AU6438KF Test ..."


'=========================================
'   POWER on
'   SD Card test
'=========================================
ChipString = "vid_058f"

If (AU6438SortingFlag) Then
    Call PowerSet2(1, "4.4", "0.3", 1, "4.4", "0.3", 1)
End If

CardResult = DO_WritePort(card, Channel_P1A, &H8E)
Call MsecDelay(0.2)

CardResult = DO_WritePort(card, Channel_P1A, &H7E)

If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If

Call MsecDelay(0.2)
rv0 = WaitDevOn(ChipString)
Call MsecDelay(0.1)

ReadLEDCounter = 0

Do
    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
    If LightOn <> 191 Then
        Call MsecDelay(0.05)
        ReadLEDCounter = ReadLEDCounter + 1
    End If
Loop Until (ReadLEDCounter > 3) Or (LightOn = 191)

If CardResult <> 0 Then
    MsgBox "Read light On fail"
    End
End If

ClosePipe
rv0 = CBWTest_New(0, 1, ChipString)
                      
If rv0 = 1 Then
    rv0 = Read_SD_Speed(0, 0, 64, "4Bits")
    ClosePipe
    If rv0 <> 1 Then
        rv0 = 2
        Tester.Print "SD bus width Fail"
    End If
End If

If rv0 = 1 Then
    rv0 = CBWTest_New_128_Sector_AU6377(0, 1)
    ClosePipe
End If

Tester.Print "rv0="; rv0


CardResult = DO_WritePort(card, Channel_P1A, &H7F)
Call MsecDelay(0.2)

If GetDeviceName(ChipString) <> "" Then
    rv0 = 0
    Tester.Print "NBMD Test Fail"
    GoTo AU6371ELResult
End If
                        
CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                      
If CardResult <> 0 Then
    MsgBox "Read light off fail"
    End
End If
                     
If rv0 = 1 Then
    If (LightOn <> &HBF) Or (LightOff <> &HFF) Then
        Tester.Print "LightON="; LightOn
        Tester.Print "LightOFF="; LightOff
        UsbSpeedTestResult = GPO_FAIL
        rv0 = 3
    End If
End If
                     
Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
If rv0 = 0 Then
    rv0 = 2 'Bin2 Only O/S fail
End If
                
'===============================================
'  CF Card test
'================================================
               
rv1 = rv0       'AU6438GIF not have CF slot
'Call LabelMenu(1, rv1, rv0)
            
'Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                

'===============================================
'  SMC Card test  : stop these test for card not enough
'================================================
              
rv2 = rv1       'AU6438GIF has no SMC slot
'Call LabelMenu(1, rv2, rv1)
                      
                
'===============================================
'  XD Card test
'================================================

rv3 = rv2       'AU6438GIF has no XD slot

'Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                
'===============================================
'  MS Card test
'================================================
                     
rv4 = rv3       'AU6344 has no MS slot pin
               
'Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'===============================================
'  MS Pro Card test
'================================================
                
rv5 = rv4       'AU6344 has no MSpro slot pin
                
'Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                 
                
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

Public Sub AU6438KFE10TestSub()
      
'2013/1/7 This code from AU6438KFF30 non support MMC card

'==================================================================
'
'  this code come from AU6438IFF30 TestSub
'
'===================================================================

Dim ChipString As String
Dim AU6371EL_SD As Byte
Dim AU6371EL_CF As Byte
Dim AU6371EL_XD As Byte
Dim AU6371EL_MS As Byte
Dim AU6371EL_MSP  As Byte
Dim AU6371EL_BootTime As Single
Dim ReadLEDCounter As Integer
Dim LoopCount As Integer

OldChipName = ""
                
'initial condition
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
                

CardResult = DO_WritePort(card, Channel_P1C, &HFF)
Call MsecDelay(0.04)

Tester.Print "AU6438KF Eng Test ..."

'=========================================
'   POWER on
'   SD Card test
'=========================================
ChipString = "vid_058f"


For LoopCount = 1 To 20
    
    If (LoopCount - 1) Mod 5 = 0 Then
        Tester.Cls
    End If
    Tester.Label3.BackColor = RGB(255, 255, 255)
    Tester.Label4.BackColor = RGB(255, 255, 255)
    Tester.Label5.BackColor = RGB(255, 255, 255)
    Tester.Label6.BackColor = RGB(255, 255, 255)
    Tester.Label7.BackColor = RGB(255, 255, 255)
    Tester.Label8.BackColor = RGB(255, 255, 255)

    rv0 = 0
    
    CardResult = DO_WritePort(card, Channel_P1A, &H8E)
    Call MsecDelay(0.2)
    
    CardResult = DO_WritePort(card, Channel_P1A, &H7E)
    
    If CardResult <> 0 Then
        MsgBox "Set SD Card Detect Down Fail"
        End
    End If
    
    Call MsecDelay(0.1)
    rv0 = WaitDevOn(ChipString)
    Call MsecDelay(0.3)
    
    ReadLEDCounter = 0
    
    Do
        CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
        If LightOn <> 191 Then
            Call MsecDelay(0.05)
            ReadLEDCounter = ReadLEDCounter + 1
        End If
    Loop Until (ReadLEDCounter > 3) Or (LightOn = 191)
    
    If CardResult <> 0 Then
        MsgBox "Read light On fail"
        End
    End If
    
    ClosePipe
    rv0 = CBWTest_New(0, 1, ChipString)
                          
    If rv0 = 1 Then
        rv0 = Read_SD_Speed(0, 0, 64, "4Bits")
        ClosePipe
        If rv0 <> 1 Then
            rv0 = 2
            Tester.Print "SD bus width Fail"
        End If
    End If
    
    If rv0 = 1 Then
        rv0 = CBWTest_New_128_Sector_AU6377(0, 1)
        ClosePipe
    End If
    
    Tester.Print "rv0="; rv0
    
    
    CardResult = DO_WritePort(card, Channel_P1A, &H7F)
    Call MsecDelay(0.3)
    
    If GetDeviceName(ChipString) <> "" Then
        rv0 = 0
        Tester.Print "NBMD Test Fail"
        'GoTo AU6371ELResult
    End If
                            
    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                          
    If CardResult <> 0 Then
        MsgBox "Read light off fail"
        End
    End If
                         
    If rv0 = 1 Then
        If (LightOn <> &HBF) Or (LightOff <> &HFF) Then
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
                   
    rv1 = rv0       'AU6438GIF not have CF slot
    'Call LabelMenu(1, rv1, rv0)
                
    'Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                    
    
    '===============================================
    '  SMC Card test  : stop these test for card not enough
    '================================================
                  
    rv2 = rv1       'AU6438GIF has no SMC slot
    'Call LabelMenu(1, rv2, rv1)
                          
                    
    '===============================================
    '  XD Card test
    '================================================
    
    rv3 = rv2       'AU6438GIF has no XD slot
    
    'Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                    
                    
    '===============================================
    '  MS Card test
    '================================================
                         
    rv4 = rv3       'AU6344 has no MS slot pin
                   
    'Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    
    
    '===============================================
    '  MS Pro Card test
    '================================================
                    
    rv5 = rv4       'AU6344 has no MSpro slot pin
                    
    'Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                 
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
    WaitDevOFF (ChipString)
    Call MsecDelay(0.1)
    
    If rv0 = 1 Then
        Tester.Print "Cycle " & LoopCount & "... PASS" & vbCrLf
    Else
        Tester.Print "Cycle " & LoopCount & "... Fail" & vbCrLf
        Exit For
    End If
    
                 
Next    'LoopCount

                
AU6371ELResult:
    
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
    WaitDevOFF (ChipString)
    Call MsecDelay(0.1)
    
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

Public Sub AU6438KFS10TestSub()
      
'2013/1/21 This code from AU6438KFS10 non support MMC card

'==================================================================
'
'  this code come from AU6438IFF30 TestSub
'
'===================================================================

Dim ChipString As String
Dim AU6371EL_SD As Byte
Dim AU6371EL_CF As Byte
Dim AU6371EL_XD As Byte
Dim AU6371EL_MS As Byte
Dim AU6371EL_MSP  As Byte
Dim AU6371EL_BootTime As Single
Dim ReadLEDCounter As Integer
Dim LoopCount As Integer
Dim k As Integer
    
OldChipName = ""
                
'initial condition
AU6371EL_SD = 1
AU6371EL_CF = 2
AU6371EL_XD = 8
AU6371EL_MS = 32
AU6371EL_MSP = 64
AU6371EL_BootTime = 0.6
                  
If PCI7248InitFinish = 0 Then
    PCI7248ExistAU6254
    Call SetTimer_1ms
    CardResult = DO_WritePort(card, Channel_P1C, &HFF)
    Call MsecDelay(0.2)
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
                

Tester.Print "AU6438KF ST1 Test ..."

'=========================================
'   POWER on
'   SD Card test
'=========================================
ChipString = "vid_058f"


For LoopCount = 1 To 1
    
    
    Tester.Label3.BackColor = RGB(255, 255, 255)
    Tester.Label4.BackColor = RGB(255, 255, 255)
    Tester.Label5.BackColor = RGB(255, 255, 255)
    Tester.Label6.BackColor = RGB(255, 255, 255)
    Tester.Label7.BackColor = RGB(255, 255, 255)
    Tester.Label8.BackColor = RGB(255, 255, 255)

    rv0 = 0
    ReaderExist = 0
    
    CardResult = DO_WritePort(card, Channel_P1A, &H7E)
    'Call MsecDelay(0.2)
    If CardResult <> 0 Then
        MsgBox "Set SD Card Detect Down Fail"
        End
    End If
    
'    Call PowerSet2(0, "3.6", "0.2", 1, "3.2", "0.25", 1)
'    Call MsecDelay(0.1)
    Call PowerSet2(1, "3.0", "0.3", 1, "3.0", "0.3", 1)
    
    rv0 = WaitDevOn(ChipString)
    
'    For k = 1 To 2
'        rv0 = CBWTest_Simple(0, ChipString)
'        If rv0 <> 1 Then
'            Debug.Print "k = "; k
'            Exit For
'        End If
'        CardResult = DO_WritePort(card, Channel_P1A, &HFF)
'        Call MsecDelay(0.2)
'        CardResult = DO_WritePort(card, Channel_P1A, &H7E)
'    Next
    
    If rv0 = 1 Then
        Call MsecDelay(0.1)
        rv0 = WaitDevOn(ChipString)
        Call MsecDelay(0.3)
    
    
        ReadLEDCounter = 0
        
        Do
            CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
            If LightOn <> 191 Then
                Call MsecDelay(0.05)
                ReadLEDCounter = ReadLEDCounter + 1
            End If
        Loop Until (ReadLEDCounter > 3) Or (LightOn = 191)
        
        If CardResult <> 0 Then
            MsgBox "Read light On fail"
            End
        End If
        
        ClosePipe
        rv0 = CBWTest_New(0, 1, ChipString)
    Else
        rv0 = 0
    End If
    
    If rv0 = 1 Then
        rv0 = Read_SD_Speed(0, 0, 64, "4Bits")
        ClosePipe
        If rv0 <> 1 Then
            rv0 = 2
            Tester.Print "SD bus width Fail"
        End If
    End If
    
    If rv0 = 1 Then
        rv0 = CBWTest_New_128_Sector_AU6377(0, 1)
        ClosePipe
    End If
    
    Tester.Print "rv0="; rv0
    
    
    CardResult = DO_WritePort(card, Channel_P1A, &H7F)
    Call MsecDelay(0.3)
    
    If GetDeviceName(ChipString) <> "" Then
        rv0 = 0
        Tester.Print "NBMD Test Fail"
        'GoTo AU6371ELResult
    End If
                            
    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                          
    If CardResult <> 0 Then
        MsgBox "Read light off fail"
        End
    End If
                         
    If rv0 = 1 Then
        If (LightOn <> &HBF) Or (LightOff <> &HFF) Then
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
                   
    rv1 = rv0       'AU6438GIF not have CF slot
    'Call LabelMenu(1, rv1, rv0)
                
    'Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                    
    
    '===============================================
    '  SMC Card test  : stop these test for card not enough
    '================================================
                  
    rv2 = rv1       'AU6438GIF has no SMC slot
    'Call LabelMenu(1, rv2, rv1)
                          
                    
    '===============================================
    '  XD Card test
    '================================================
    
    rv3 = rv2       'AU6438GIF has no XD slot
    
    'Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                    
                    
    '===============================================
    '  MS Card test
    '================================================
                         
    rv4 = rv3       'AU6344 has no MS slot pin
                   
    'Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    
    
    '===============================================
    '  MS Pro Card test
    '================================================
                    
    rv5 = rv4       'AU6344 has no MSpro slot pin
                    
    'Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                 
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
    WaitDevOFF (ChipString)
    Call MsecDelay(0.1)
    
'    If rv0 = 1 Then
'        Tester.Print "Cycle " & LoopCount & "... PASS" & vbCrLf
'    Else
'        Tester.Print "Cycle " & LoopCount & "... Fail" & vbCrLf
'        Exit For
'    End If
    
    SetSiteStatus (RunHV)
    Call WaitAnotherSiteDone(RunHV, 3)
    Call PowerSet2(1, "0.0", "0.3", 1, "0.0", "0.3", 1)
'    Call MsecDelay(0.2)
                 
Next    'LoopCount

                
AU6371ELResult:
    
    SetSiteStatus (UNKNOW)
    
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

Public Sub AU6438KFS00TestSub()

'==================================================================
'
'  this code come from AU6438IFF00TestSub
'  2012/11/13 HV + LV test, using FT6 socket board
'
'===================================================================
'2013/7/17 Copy from AU6438KFF00TestSub, only modify bus width
'2014/3/25 Add sorting condition (36ohm)

    Dim ChipString As String
    Dim AU6371EL_SD As Byte
    Dim AU6371EL_CF As Byte
    Dim AU6371EL_XD As Byte
    Dim AU6371EL_MS As Byte
    Dim AU6371EL_MSP  As Byte
    Dim AU6371EL_BootTime As Single
    Dim HV_Done_Flag As Boolean
    Dim HV_Result As String
    Dim LV_Result As String
    OldChipName = ""
                 
    ' initial condition
                
    AU6371EL_SD = 1
    AU6371EL_CF = 2
    AU6371EL_XD = 8
    AU6371EL_MS = 32
    AU6371EL_MSP = 64
    AU6371EL_BootTime = 0.6
    
    HV_Flag = False
    HV_Result = 0
    LV_Result = 0
                     
    If PCI7248InitFinish_Sync = 0 Then
        PCI7248Exist_P1C_Sync
    End If
               
    CardResult = DO_WritePort(card, Channel_P1C, &HFF)
    Call MsecDelay(0.2)
    ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
    '  CardResult = DO_WritePort(card, Channel_P1B, &H0)

Routine_Label:


    LBA = LBA + 1
                         
    If Not HV_Done_Flag Then
        Call PowerSet2(0, "5.25", "0.3", 1, "5.25", "0.3", 1)
        Call MsecDelay(0.2)
        Tester.Print "5.25V Begin Test ..."
        SetSiteStatus (RunHV)
    Else
        Call PowerSet2(0, "4.4", "0.3", 1, "4.4", "0.3", 1)
        Call MsecDelay(0.2)
        Tester.Print vbCrLf & "4.4V Begin Test ..."
        SetSiteStatus (RunLV)
    End If

                         
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

    CardResult = DO_WritePort(card, Channel_P1A, &H8E)
    Call MsecDelay(0.3)
    
    CardResult = DO_WritePort(card, Channel_P1A, &H7E)
    
    If CardResult <> 0 Then
        MsgBox "Set SD Card Detect Down Fail"
        End
    End If
    
    Call MsecDelay(0.1)
    rv0 = WaitDevOn(ChipString)
    Call MsecDelay(0.3)
    
    ReadLEDCounter = 0
    
    Do
        CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
        If LightOn <> 191 Then
            Call MsecDelay(0.05)
            ReadLEDCounter = ReadLEDCounter + 1
        End If
    Loop Until (ReadLEDCounter > 3) Or (LightOn = 191)
    
    If CardResult <> 0 Then
        MsgBox "Read light On fail"
        End
    End If
    
    ClosePipe
    rv0 = CBWTest_New(0, 1, ChipString)
                          
    If rv0 = 1 Then
        rv0 = Read_SD_Speed(0, 0, 64, "4Bits")
        ClosePipe
        If rv0 <> 1 Then
            rv0 = 2
            Tester.Print "SD bus width Fail"
        End If
    End If
    
    If rv0 = 1 Then
        rv0 = CBWTest_New_128_Sector_AU6377(0, 1)
        ClosePipe
    End If
    
    Tester.Print "rv0="; rv0
    
    
    CardResult = DO_WritePort(card, Channel_P1A, &H7F)
    Call MsecDelay(0.3)
    
    If GetDeviceName(ChipString) <> "" Then
        rv0 = 0
        Tester.Print "NBMD Test Fail"
        'GoTo AU6371ELResult
    End If
                            
    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                          
    If CardResult <> 0 Then
        MsgBox "Read light off fail"
        End
    End If
                         
    If rv0 = 1 Then
        If (LightOn <> &HBF) Or (LightOff <> &HFF) Then
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
               
    rv1 = rv0  '----------- AU6438CF has no CF slot
                 
    'Call LabelMenu(1, rv1, rv0)
    'Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    
    '===============================================
    '  SMC Card test  : stop these test for card not enough
    '================================================
              
    'AU6438 has no SMC slot
               
    rv2 = rv1   ' to complete the SMC asbolish
               
              
    '===============================================
    '  XD Card test
    '================================================
    
    rv3 = rv2
                
    '===============================================
    '  MS Card test
    '================================================
                
    rv4 = rv3  'AU6344 has no MS slot pin
               
    'Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    '===============================================
    '  MS Pro Card test
    '================================================
              
    rv5 = rv4 'AU6438JCF has no MSpro slot
                  
                  
    '=================================================
    '   NO Card test
    '=================================================
                
'    CardResult = DO_WritePort(card, Channel_P1A, &H7F)   ' Close power
'    Call MsecDelay(0.2)
'
'    If GetDeviceName(ChipString) = "" Then
'        rv0 = 0
'    End If
                
                
AU6371ELResult:
                
    ClosePipe
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
    WaitDevOFF ("pid_6366")
    SetSiteStatus (SiteUnknow)
    
    If HV_Done_Flag = False Then
        If rv0 <> 1 Then
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
        Call MsecDelay(0.4)
        ReaderExist = 0
        GoTo Routine_Label
    Else
        If rv0 <> 1 Then
            LV_Result = "Bin2"
            Tester.Print "LV Unknow"
        ElseIf rv0 * rv1 * rv2 * rv3 <> 1 Then
            LV_Result = "Fail"
            Tester.Print "LV Fail"
        ElseIf rv0 * rv1 * rv2 * rv3 = 1 Then
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
                
End Sub


Public Sub AU6438IFS00TestSub()


'==================================================================
'
'  this code come from AU6438IFF00TestSub
'  2012/11/13 HV + LV test, using FT6 socket board
'
'===================================================================
'2013/7/17 Copy from AU6438KFF00TestSub, only modify bus width
'2014/3/25 Add sorting condition (36ohm)

    Dim ChipString As String
    Dim AU6371EL_SD As Byte
    Dim AU6371EL_CF As Byte
    Dim AU6371EL_XD As Byte
    Dim AU6371EL_MS As Byte
    Dim AU6371EL_MSP  As Byte
    Dim AU6371EL_BootTime As Single
    Dim HV_Done_Flag As Boolean
    Dim HV_Result As String
    Dim LV_Result As String
    OldChipName = ""
                 
    ' initial condition
                
    AU6371EL_SD = 1
    AU6371EL_CF = 2
    AU6371EL_XD = 8
    AU6371EL_MS = 32
    AU6371EL_MSP = 64
    AU6371EL_BootTime = 0.6
    
    HV_Flag = False
    HV_Result = 0
    LV_Result = 0
                     
    If PCI7248InitFinish_Sync = 0 Then
        PCI7248Exist_P1C_Sync
    End If
               
    CardResult = DO_WritePort(card, Channel_P1C, &HFF)
    Call MsecDelay(0.2)
    ' result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
    '  CardResult = DO_WritePort(card, Channel_P1B, &H0)

Routine_Label:


    LBA = LBA + 1
                         
    If Not HV_Done_Flag Then
        Call PowerSet2(0, "5.25", "0.3", 1, "5.25", "0.3", 1)
        Call MsecDelay(0.2)
        Tester.Print "5.25V Begin Test ..."
        SetSiteStatus (RunHV)
    Else
        Call PowerSet2(0, "4.4", "0.3", 1, "4.4", "0.3", 1)
        Call MsecDelay(0.2)
        Tester.Print vbCrLf & "4.4V Begin Test ..."
        SetSiteStatus (RunLV)
    End If

                         
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

    CardResult = DO_WritePort(card, Channel_P1A, &H8E)
    Call MsecDelay(0.3)
    
    CardResult = DO_WritePort(card, Channel_P1A, &H7E)
    
    If CardResult <> 0 Then
        MsgBox "Set SD Card Detect Down Fail"
        End
    End If
    
    Call MsecDelay(0.1)
    rv0 = WaitDevOn(ChipString)
    Call MsecDelay(0.3)
    
    ReadLEDCounter = 0
    
    Do
        CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
        If LightOn <> 191 Then
            Call MsecDelay(0.05)
            ReadLEDCounter = ReadLEDCounter + 1
        End If
    Loop Until (ReadLEDCounter > 3) Or (LightOn = 191)
    
    If CardResult <> 0 Then
        MsgBox "Read light On fail"
        End
    End If
    
    ClosePipe
    rv0 = CBWTest_New(0, 1, ChipString)
                          
    If rv0 = 1 Then
        rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
        ClosePipe
        If rv0 <> 1 Then
            rv0 = 2
            Tester.Print "SD bus width Fail"
        End If
    End If
    
    If rv0 = 1 Then
        rv0 = CBWTest_New_128_Sector_AU6377(0, 1)
        ClosePipe
    End If
    
    Tester.Print "rv0="; rv0
    
    
    CardResult = DO_WritePort(card, Channel_P1A, &H7F)
    Call MsecDelay(0.3)
    
    If GetDeviceName(ChipString) <> "" Then
        rv0 = 0
        Tester.Print "NBMD Test Fail"
        'GoTo AU6371ELResult
    End If
                            
    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                          
    If CardResult <> 0 Then
        MsgBox "Read light off fail"
        End
    End If
                         
    If rv0 = 1 Then
        If (LightOn <> &HBF) Or (LightOff <> &HFF) Then
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
               
    rv1 = rv0  '----------- AU6438CF has no CF slot
                 
    'Call LabelMenu(1, rv1, rv0)
    'Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    
    '===============================================
    '  SMC Card test  : stop these test for card not enough
    '================================================
              
    'AU6438 has no SMC slot
               
    rv2 = rv1   ' to complete the SMC asbolish
               
              
    '===============================================
    '  XD Card test
    '================================================
    
    rv3 = rv2
                
    '===============================================
    '  MS Card test
    '================================================
                
    rv4 = rv3  'AU6344 has no MS slot pin
               
    'Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    '===============================================
    '  MS Pro Card test
    '================================================
              
    rv5 = rv4 'AU6438JCF has no MSpro slot
                  
                  
    '=================================================
    '   NO Card test
    '=================================================
                
'    CardResult = DO_WritePort(card, Channel_P1A, &H7F)   ' Close power
'    Call MsecDelay(0.2)
'
'    If GetDeviceName(ChipString) = "" Then
'        rv0 = 0
'    End If
                
                
AU6371ELResult:
                
    ClosePipe
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
    WaitDevOFF ("pid_6366")
    SetSiteStatus (SiteUnknow)
    
    If HV_Done_Flag = False Then
        If rv0 <> 1 Then
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
        Call MsecDelay(0.4)
        ReaderExist = 0
        GoTo Routine_Label
    Else
        If rv0 <> 1 Then
            LV_Result = "Bin2"
            Tester.Print "LV Unknow"
        ElseIf rv0 * rv1 * rv2 * rv3 <> 1 Then
            LV_Result = "Fail"
            Tester.Print "LV Fail"
        ElseIf rv0 * rv1 * rv2 * rv3 = 1 Then
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
                
                
End Sub

Public Sub AU6438IFF01TestSub()
      
'2011/7/15 add FT6
'2011/10/03 Using socket-board "AU6431-GEF 28QFN(OS FT) SOCKET V1.00"
'20121/1/13 Add for FT3 modified by FT6(AU6438IFF31) code
'           socket-board "AU6431-GEF 28QFN(OS FT) SOCKET V1.00" (Remove R40)

Dim ChipString As String
Dim ReadLEDCounter As Integer
Dim HV_Flag As Boolean

HV_Flag = False

If PCI7248InitFinish = 0 Then
    PCI7248ExistAU6254
    Call SetTimer_1ms
End If
               
'==========================================================
'    Start Open Shot test
'==========================================================

CardResult = DO_WritePort(card, Channel_P1C, &HFF)

'=========================================
'   POWER on
'   SD Card test
'=========================================
ChipString = "vid_058f"

HVLV_Label:

LBA = LBA + 1
LightOn = 0
                         
rv0 = 3     'Enum
rv1 = 3     'SD
rv2 = 3     'SD R/W 64K
rv3 = 3     'SD Speed
rv4 = 3     'NBMD
rv5 = 3     'LED
             
Tester.Label3.BackColor = RGB(255, 255, 255)
Tester.Label4.BackColor = RGB(255, 255, 255)
Tester.Label5.BackColor = RGB(255, 255, 255)
Tester.Label6.BackColor = RGB(255, 255, 255)
Tester.Label7.BackColor = RGB(255, 255, 255)
Tester.Label8.BackColor = RGB(255, 255, 255)


If (HV_Flag = False) Then
    Call PowerSet2(1, "3.6", "0.5", 1, "3.6", "0.5", 1)
    Tester.Print "Begin HV 3.6 Test ..."
Else
    Call PowerSet2(1, "3.0", "0.5", 1, "3.0", "0.5", 1)
    Tester.Print vbCrLf & "Begin LV 3.0 Test ..."
    Call MsecDelay(0.2)
End If

CardResult = DO_WritePort(card, Channel_P1A, &H8E)
Call MsecDelay(0.2)

CardResult = DO_WritePort(card, Channel_P1A, &H7E)

If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If

Call MsecDelay(0.2)
rv0 = WaitDevOn(ChipString)
Call MsecDelay(0.1)
Call NewLabelMenu(0, "Enum", rv0, 1)

ReadLEDCounter = 0
Do
    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
    If LightOn <> 191 Then
        Call MsecDelay(0.05)
        ReadLEDCounter = ReadLEDCounter + 1
    End If
Loop Until (ReadLEDCounter > 3) Or (LightOn = 191)


ClosePipe
rv1 = CBWTest_New(0, 1, ChipString)
Call NewLabelMenu(1, "SD", rv1, rv0)
     
rv2 = CBWTest_New_128_Sector_PipeReady(0, rv1)
Call NewLabelMenu(2, "SD 64K", rv2, rv1)

If rv2 = 1 Then
    rv3 = Read_SD_Speed(LBA, 0, 64, "4Bits")
    ClosePipe
    If rv3 <> 1 Then
        rv3 = 2
        Tester.Print "SD bus width Fail"
    End If
    Call NewLabelMenu(3, "SD Bus Width/Speed", rv3, rv2)
End If

ClosePipe

CardResult = DO_WritePort(card, Channel_P1A, &H7F)
Call MsecDelay(0.2)

If rv3 = 1 Then
    rv4 = WaitDevOFF(ChipString)
    Call NewLabelMenu(4, "NB Mode", rv4, rv3)
                            
                            
    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                          
    If CardResult <> 0 Then
        MsgBox "Read light off fail"
        End
    End If
                         
    If rv4 = 1 Then
        If (LightOn = &HBF) And (LightOff = &HFF) Then
            rv5 = 1
        Else
            Tester.Print "LightON="; LightOn
            Tester.Print "LightOFF="; LightOff
            UsbSpeedTestResult = GPO_FAIL
            rv5 = 3
        End If
        Call NewLabelMenu(5, "GPO", rv5, rv4)
    End If
              
End If

AU6438Result:
    
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
    Call PowerSet2(1, "0.0", "0.2", 1, "0.0", "0.2", 1)
    
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
            
            ReaderExist = 0
            HV_Flag = True
            Call MsecDelay(0.2)
            GoTo HVLV_Label
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
