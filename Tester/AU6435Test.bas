Attribute VB_Name = "AU6435Test"
Public Sub AU6435TestSub()

    If ChipName = "AU6435DLF20" Then
        Call AU6435DLF20TestSub
    ElseIf ChipName = "AU6435DLF21" Then
        Call AU6435DLF21TestSub
    ElseIf ChipName = "AU6435DLF22" Then
        Call AU6435DLF22TestSub
    ElseIf ChipName = "AU6435DLF23" Then
        Call AU6435DLF23TestSub
    ElseIf ChipName = "AU6435DLF24" Then
        Call AU6435DLF24TestSub
    ElseIf ChipName = "AU6435ELF04" Then
        Call AU6435ELF04TestSub
    ElseIf ChipName = "AU6435ELF21" Then
        Call AU6435ELF21TestSub
    ElseIf ChipName = "AU6435ELF22" Then
        Call AU6435ELF22TestSub
    ElseIf ChipName = "AU6435ELF23" Then
        Call AU6435ELF23TestSub
    ElseIf ChipName = "AU6435ELF24" Then
        Call AU6435ELF24TestSub
    ElseIf ChipName = "AU6435ELF25" Then
        Call AU6435ELF25TestSub
    ElseIf ChipName = "AU6435GLE10" Then
        Call AU6435GLE10TestSub
    ElseIf ChipName = "AU6435ELF33" Then
        Call AU6435ELF33TestSub
    ElseIf ChipName = "AU6435ELF34" Then
        Call AU6435ELF34TestSub
    ElseIf ChipName = "AU6435DLF2D" Then
        Call AU6435DLF2DTestSub
    ElseIf ChipName = "AU6435ELS11" Then
        Call AU6435ELS11TestSub
    ElseIf ChipName = "AU6435BLF21" Then
        Call AU6435BLF21TestSub
    ElseIf ChipName = "AU6435AFF21" Then
        Call AU6435AFF21TestSub
    ElseIf ChipName = "AU6435BFF20" Then
        Call AU6435BFF20TestSub
    ElseIf ChipName = "AU6435BFF21" Then
        Call AU6435BFF21TestSub
    ElseIf ChipName = "AU6435CFF22" Then
        Call AU6435CFF22TestSub
    ElseIf ChipName = "AU6435BFF23" Then
        Call AU6435BFF23TestSub
    ElseIf ChipName = "AU6435BFS11" Then
        Call AU6435BFS11TestSub
    End If
    
End Sub


Public Sub AU6435DLF20TestSub()

'insert SD 2.0 Memory Card
Dim TmpLBA As Long
Dim i As Integer
'Call PowerSet2(1, "3.3", "0.05", 1, "2.1", "0.05", 1)
      
Tester.Print "AU6435DL : Begin Test ..."
'==================================================================
'
'  this code come from AU6433DLF20TestSub
'  Add HID test function
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
                  
                Call MsecDelay(0.3)    'power on time
                
                ChipString = "vid"
                 
                If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6435DLFResult
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
                CardResult = DO_WritePort(card, Channel_P1A, &H3E)  'use 48MHz clock source
                      
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
                      
                'ClosePipe
                      
                rv0 = CBWTest_New(0, 1, ChipString)
                      
                If rv0 = 1 Then
                    rv0 = Read_SD_Speed_AU6435(0, 0, 64, "8Bits")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
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
                    GoTo AU6435DLFResult
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
                CardResult = DO_WritePort(card, Channel_P1A, &H36) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                End If
                  
                  
                Call MsecDelay(0.1)
                
                CardResult = DO_WritePort(card, Channel_P1A, &H37)  'XD
                
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
              
                
                
                CardResult = DO_WritePort(card, Channel_P1A, &H37)      'MS + XD
              
                Call MsecDelay(0.1)
               
                CardResult = DO_WritePort(card, Channel_P1A, &H1F)      'MS
               
                 
                Call MsecDelay(0.1)
                
                OpenPipe
                rv5 = ReInitial(0)
                ClosePipe
                
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                
                
                    If rv5 = 1 Then
                        rv5 = Read_MS_Speed_AU6435(0, 0, 64, "4Bits")
                        
                        If rv5 <> 1 Then
                            rv5 = 2
                            Tester.Print "MS bus width Fail"
                        End If
                    End If
                
                Call LabelMenu(31, rv5, rv4)
                
                Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                ClosePipe
                
                If rv5 <> 1 Then
                    GoTo AU6435DLFResult
                End If
                
                                
                '=================================================================================
                ' HID mode and reader mode ---> compositive device
                If rv5 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &HAF) '  pwr off  for HID mode
                    'result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                    
                    Call MsecDelay(0.2)
          
                    CardResult = DO_WritePort(card, Channel_P1A, &H6E) ' HID mode   'PID_6435
                            
                    Call MsecDelay(1.3)
                    
                    LightOn = 0
                    
                    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
          
                    If rv5 = 1 And (LightOn <> 127 And LightOff <> 255) Then
                        UsbSpeedTestResult = GPO_FAIL
                        rv4 = 2
                        Tester.Label9.Caption = "GPO FAIL " & LightOff
                    End If
                
                End If
          
                If rv4 = 1 Then
                ' code begin
                     
                    Tester.Cls
                    Tester.Print "keypress test begin---------------"
                    Dim ReturnValue As Byte
                     
                    DeviceHandle = &HFFFF  'invalid handle initial value
                     
                    ReturnValue = fnGetDeviceHandle(DeviceHandle)
                    Tester.Print ReturnValue; Space(5); ' 1: pass the other refer btnstatus.h
                    Tester.Print "DeviceHandle="; DevicehHandle
                     
                    If ReturnValue <> 1 Then
                        rv0 = UNKNOW       '---> HID mode unknow device mode
                        Call LabelMenu(0, rv0, 1)
                        Tester.Label9.Caption = "HID mode unknow device"
                        fnFreeDeviceHandle (DeviceHandle)
                        GoTo AU6435DLFResult
                    End If
                     
                    '=======================
                    '  key press test, it will return 10 when key up, GPI 6 must do low go hi action
                    '========================
                     
                
                    Do
                        CardResult = DO_WritePort(card, Channel_P1A, &H6E) 'GPI6 : bit 7: pull high
                        Sleep (200)
                        CardResult = DO_WritePort(card, Channel_P1A, &H2E)  ' GPI6 : bit 7: pull low
                        Sleep (1000)
                       
                        ReturnValue = fnInquiryBtnStatus(DeviceHandle)
                        Tester.Print i; Space(5); "Key press value="; ReturnValue
                        i = i + 1
                    Loop While i < 3 And ReturnValue <> 10
                     
                    If ReturnValue <> 10 Then
                     
                        rv1 = 2
                        Call LabelMenu(1, rv1, rv0)
                        Label9.Caption = "KeyPress Fail"
                       
                    End If
                              
                End If
                
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print "LBA="; LBA
                 
               
AU6435DLFResult:
                
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
Public Sub AU6435DLF21TestSub()

'insert SD 3.0 Memory Card
Dim TmpLBA As Long
Dim i As Integer
'Call PowerSet2(1, "3.3", "0.05", 1, "2.1", "0.05", 1)
      
Tester.Print "AU6435DL : Begin Test ..."
'==================================================================
'
'  this code come from AU6433DLF20TestSub
'  Add HID test function
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
               
               '  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
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
                  
                Call MsecDelay(0.3)    'power on time
                
                ChipString = "vid"
                 
                'If GetDeviceName(ChipString) <> "" Then
                '    rv0 = 0
                '    GoTo AU6435DLFResult
                'End If
                 
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
                CardResult = DO_WritePort(card, Channel_P1A, &H3E)
                      
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
                      
                'ClosePipe
                      
                rv0 = CBWTest_New(0, 1, ChipString)
                      
                If rv0 = 1 Then
                    rv0 = Read_SD30_Speed_AU6435(0, 0, 64, "4Bits")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
                    End If
                End If
                
                If rv0 = 1 Then
                    rv0 = Read_SD30_Mode_AU6435(0, 0, 64, "SDR")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD3.0 Mode Fail"
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
                    GoTo AU6435DLFResult
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
                CardResult = DO_WritePort(card, Channel_P1A, &H36) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                End If
                  
                  
                Call MsecDelay(0.1)
                
                CardResult = DO_WritePort(card, Channel_P1A, &H37)  'XD
                
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
              
                
                
                CardResult = DO_WritePort(card, Channel_P1A, &H37)      'MS + XD
              
                Call MsecDelay(0.1)
               
                CardResult = DO_WritePort(card, Channel_P1A, &H1F)      'MS
               
                 
                Call MsecDelay(0.1)
                
                OpenPipe
                rv5 = ReInitial(0)
                ClosePipe
                
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                
                
                    If rv5 = 1 Then
                        rv5 = Read_MS_Speed_AU6435(0, 0, 64, "4Bits")
                        
                        If rv5 <> 1 Then
                            rv5 = 2
                            Tester.Print "MS bus width Fail"
                        End If
                    End If
                
                Call LabelMenu(31, rv5, rv4)
                
                Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                ClosePipe
                
                If rv5 <> 1 Then
                    GoTo AU6435DLFResult
                End If
                
                                
                '=================================================================================
                ' HID mode and reader mode ---> compositive device
                If rv5 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &HAF) '  pwr off  for HID mode
                    'result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                    
                    Call MsecDelay(0.2)
          
                    CardResult = DO_WritePort(card, Channel_P1A, &H6E) ' HID mode   'PID_6466
                            
                    Call MsecDelay(1.3)
                    
                    LightOn = 0
                    
                    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
          
                    If rv5 = 1 And (LightOn <> 127 And LightOff <> 255) Then
                        UsbSpeedTestResult = GPO_FAIL
                        rv4 = 2
                        Tester.Label9.Caption = "GPO FAIL " & LightOff
                    End If
                
                End If
          
                If rv4 = 1 Then
                ' code begin
                     
                    Tester.Cls
                    Tester.Print "keypress test begin---------------"
                    Dim ReturnValue As Byte
                     
                    DeviceHandle = &HFFFF  'invalid handle initial value
                     
                    ReturnValue = fnGetDeviceHandle(DeviceHandle)
                    Tester.Print ReturnValue; Space(5); ' 1: pass the other refer btnstatus.h
                    Tester.Print "DeviceHandle="; DevicehHandle
                     
                    If ReturnValue <> 1 Then
                        rv0 = UNKNOW       '---> HID mode unknow device mode
                        Call LabelMenu(0, rv0, 1)
                        Tester.Label9.Caption = "HID mode unknow device"
                        fnFreeDeviceHandle (DeviceHandle)
                        GoTo AU6435DLFResult
                    End If
                     
                    '=======================
                    '  key press test, it will return 10 when key up, GPI 6 must do low go hi action
                    '========================
                     
                
                    Do
                        CardResult = DO_WritePort(card, Channel_P1A, &H6E) 'GPI6 : bit 6: pull high
                        Sleep (200)
                        CardResult = DO_WritePort(card, Channel_P1A, &H2E)  ' GPI6 : bit 6: pull low
                        Sleep (1000)
                       
                        ReturnValue = fnInquiryBtnStatus(DeviceHandle)
                        Tester.Print i; Space(5); "Key press value="; ReturnValue
                        i = i + 1
                    Loop While i < 3 And ReturnValue <> 10
                     
                    If ReturnValue <> 10 Then
                     
                        rv1 = 2
                        Call LabelMenu(1, rv1, rv0)
                        Label9.Caption = "KeyPress Fail"
                       
                    End If
                              
                End If
                
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print "LBA="; LBA
                 
               
AU6435DLFResult:
                
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

Public Sub AU6435DLF22TestSub()

'insert SD 3.0 Memory Card
'99/12/01 add V18-out detect
Dim TmpLBA As Long
Dim i As Integer
Dim Detect_Counter As Integer
'Call PowerSet2(1, "1.87", "0.05", 1, "1.53", "0.05", 1)
      
Tester.Print "AU6435DL : Begin Test ..."
'==================================================================
'
'  this code come from AU6433DLF20TestSub
'  Add HID test function
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
               
               '  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
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
                  
                Call MsecDelay(0.2)    'power on time
                
                ChipString = "vid"
                 
                If GetDeviceName_NoReply(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6435DLFResult
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
                CardResult = DO_WritePort(card, Channel_P1A, &H3E)
                      
                If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect Down Fail"
                    End
                End If
                
                Call MsecDelay(0.5)
                
                Do
                    If Trim(GetDeviceName_NoReply("vid")) <> "" Then
                        Detect_Counter = 15
                    End If
                    Call MsecDelay(0.1)
                    Detect_Counter = Detect_Counter + 1
                
                Loop Until (Detect_Counter > 14)
                Detect_Counter = 0
                Call MsecDelay(0.1)
                
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                
                If CardResult <> 0 Then
                    MsgBox "Read light On fail"
                    End
                End If
                      
                'ClosePipe
                      
                rv0 = CBWTest_New(0, 1, ChipString)
                      
                If rv0 = 1 Then
                    rv0 = Read_SD30_Speed_AU6435(0, 0, 64, "4Bits")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
                    End If
                End If
                
                If rv0 = 1 Then
                    rv0 = Read_SD30_Mode_AU6435(0, 0, 64, "SDR")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD3.0 Mode Fail"
                    End If
                End If
                
                
                ClosePipe
                      
                Tester.Print "rv0="; rv0
                     
                If rv0 <> 0 Then
                    If LightOn <> &H1F Or LightOff <> &H9F Then
                        Tester.Print "LightON="; LightOn
                        Tester.Print "LightOFF="; LightOff
                        Tester.Print "GPO_Fail or V18 out of range"
                        UsbSpeedTestResult = GPO_FAIL
                        rv0 = 3
                    Else
                        Tester.Print "V18-out Range PASS"
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
                    GoTo AU6435DLFResult
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
                CardResult = DO_WritePort(card, Channel_P1A, &H36) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                End If
                  
                  
                Call MsecDelay(0.1)
                
                CardResult = DO_WritePort(card, Channel_P1A, &H37)  'XD
                
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
              
                
                
                CardResult = DO_WritePort(card, Channel_P1A, &H37)      'MS + XD
              
                Call MsecDelay(0.1)
               
                CardResult = DO_WritePort(card, Channel_P1A, &H1F)      'MS
               
                 
                Call MsecDelay(0.1)
                
                OpenPipe
                rv5 = ReInitial(0)
                ClosePipe
                
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                
                
                    If rv5 = 1 Then
                        rv5 = Read_MS_Speed_AU6435(0, 0, 64, "4Bits")
                        
                        If rv5 <> 1 Then
                            rv5 = 2
                            Tester.Print "MS bus width Fail"
                        End If
                    End If
                
                Call LabelMenu(31, rv5, rv4)
                
                Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                ClosePipe
                
                If rv5 <> 1 Then
                    GoTo AU6435DLFResult
                End If
                
                                
                '=================================================================================
                ' HID mode and reader mode ---> compositive device
                If rv5 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &HAF) '  pwr off  for HID mode
                    'result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                    
                    Call MsecDelay(0.2)
          
                    CardResult = DO_WritePort(card, Channel_P1A, &H6E) ' HID mode   'PID_6466
                            
                    Call MsecDelay(0.5)
                    
                    Do
                        If Trim(GetDeviceName_NoReply("vid")) <> "" Then
                            Detect_Counter = 15
                        End If
                        Call MsecDelay(0.1)
                        Detect_Counter = Detect_Counter + 1
                
                    Loop Until (Detect_Counter > 14)
                    Detect_Counter = 0
                    Call MsecDelay(0.1)
                    
                    LightOn = 0
                    
                    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
          
                    If rv5 = 1 And (LightOn <> &H1F Or LightOff <> &H9F) Then
                        UsbSpeedTestResult = GPO_FAIL
                        rv4 = 2
                        Tester.Label9.Caption = "GPO FAIL " & LightOff
                    End If
                
                End If
          
                If rv4 = 1 Then
                ' code begin
                     
                    Tester.Cls
                    Tester.Print "keypress test begin---------------"
                    Dim ReturnValue As Byte
                     
                    DeviceHandle = &HFFFF  'invalid handle initial value
                     
                    ReturnValue = fnGetDeviceHandle(DeviceHandle)
                    Tester.Print ReturnValue; Space(5); ' 1: pass the other refer btnstatus.h
                    Tester.Print "DeviceHandle="; DevicehHandle
                     
                    If ReturnValue <> 1 Then
                        rv0 = UNKNOW       '---> HID mode unknow device mode
                        Call LabelMenu(0, rv0, 1)
                        Tester.Label9.Caption = "HID mode unknow device"
                        fnFreeDeviceHandle (DeviceHandle)
                        GoTo AU6435DLFResult
                    End If
                     
                    '=======================
                    '  key press test, it will return 10 when key up, GPI 6 must do low go hi action
                    '========================
                     
                
                    Do
                        CardResult = DO_WritePort(card, Channel_P1A, &H6E) 'GPI6 : bit 6: pull high
                        Sleep (200)
                        CardResult = DO_WritePort(card, Channel_P1A, &H2E)  ' GPI6 : bit 6: pull low
                        Sleep (1000)
                       
                        ReturnValue = fnInquiryBtnStatus(DeviceHandle)
                        Tester.Print i; Space(5); "Key press value="; ReturnValue
                        i = i + 1
                    Loop While i < 3 And ReturnValue <> 10
                     
                    If ReturnValue <> 10 Then
                     
                        rv1 = 2
                        Call LabelMenu(1, rv1, rv0)
                        Label9.Caption = "KeyPress Fail"
                       
                    End If
                              
                End If
                
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print "LBA="; LBA
                 
               
AU6435DLFResult:
                
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

Public Sub AU6435BLF21TestSub()

'insert SD 3.0 Memory Card
'99/12/01 add V18-out detect
Dim TmpLBA As Long
Dim i As Integer
Dim Detect_Counter As Integer
'Call PowerSet2(1, "1.87", "0.05", 1, "1.53", "0.05", 1)
      
Tester.Print "AU6435DL : Begin Test ..."
'==================================================================
'
'  this code come from AU6433DLF20TestSub
'  Add HID test function
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
               
               '  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
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
                
                If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                End If
                
                'Call MsecDelay(0.2)
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                Call MsecDelay(0.7)     'power on time
                
                ChipString = "vid"
                 
                
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
                CardResult = DO_WritePort(card, Channel_P1A, &H3E)
                      
                If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect Down Fail"
                    End
                End If
                
                Call MsecDelay(0.3)
                
                rv0 = WaitDevOn(ChipString)
                
                Call MsecDelay(0.2)
                
                OpenPipe
                rv0 = AU6435Close_OverCurrent(rv0)
                ClosePipe
                
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                
                If CardResult <> 0 Then
                    MsgBox "Read light On fail"
                    End
                End If
                      
                'ClosePipe
                If rv0 = 1 Then
                    rv0 = AU6435_CBWTest_New(0, 1, ChipString)
                End If
                
                
                
                If rv0 = 1 Then
                    rv0 = Read_SD30_Speed_AU6435(0, 0, 64, "4Bits")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
                    End If
                End If
                
                If rv0 = 1 Then
                    rv0 = Read_SD30_Mode_AU6435(0, 0, 64, "SDR")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD3.0 Mode Fail"
                    End If
                End If
                
                
                ClosePipe
                      
                Tester.Print "rv0="; rv0
                     
                If rv0 <> 0 Then
                    If LightOn <> &H13 Or LightOff <> &H93 Then
                        Tester.Print "LightON="; LightOn
                        Tester.Print "LightOFF="; LightOff
                        Tester.Print "GPO_Fail or V18 out of range"
                        UsbSpeedTestResult = GPO_FAIL
                        rv0 = 3
                    Else
                        Tester.Print "V18-out Range PASS"
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
                
                If rv0 = 1 Then
                    rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                End If
                                
                If rv1 <> 1 Then
                    LBA = TmpLBA
                    GoTo AU6435DLFResult
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
                CardResult = DO_WritePort(card, Channel_P1A, &H36) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                End If
                  
                  
                Call MsecDelay(0.1)
                
                CardResult = DO_WritePort(card, Channel_P1A, &H37)  'XD
                
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
              
                
                
                CardResult = DO_WritePort(card, Channel_P1A, &H37)      'MS + XD
              
                Call MsecDelay(0.1)
               
                CardResult = DO_WritePort(card, Channel_P1A, &H1F)      'MS
               
                 
                Call MsecDelay(0.1)
                
                OpenPipe
                rv5 = ReInitial(0)
                ClosePipe
                
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                
                
                    If rv5 = 1 Then
                        rv5 = Read_MS_Speed_AU6435(0, 0, 64, "4Bits")
                        
                        If rv5 <> 1 Then
                            rv5 = 2
                            Tester.Print "MS bus width Fail"
                        End If
                    End If
                
                Call LabelMenu(31, rv5, rv4)
                
                Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                ClosePipe
                
                If rv5 <> 1 Then
                    GoTo AU6435DLFResult
                End If
                
                CardResult = DO_WritePort(card, Channel_P1A, &H3F)
                Call MsecDelay(0.2)
                
                If GetDeviceName_NoReply(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6435DLFResult
                End If
                                
                '=================================================================================
                ' HID mode and reader mode ---> compositive device
                If rv5 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &HAF) '  pwr off  for HID mode
                    'result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                    
                    
                    
                    Call MsecDelay(0.2)
          
                    CardResult = DO_WritePort(card, Channel_P1A, &H6E) ' HID mode   'PID_6466
                            
                    Call MsecDelay(0.3)
                    
                    rv5 = WaitDevOn(ChipString)
                    Call MsecDelay(0.2)
                    Detect_Counter = 0
                    
                    LightOn = 0
                    
                    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
          
                    If rv5 = 1 And (LightOn <> &H13 Or LightOff <> &H93) Then
                        UsbSpeedTestResult = GPO_FAIL
                        rv4 = 2
                        Tester.Label9.Caption = "GPO FAIL " & LightOff
                    End If
                
                End If
          
                If rv4 = 1 Then
                ' code begin
                     
                    Tester.Cls
                    Tester.Print "keypress test begin---------------"
                    Dim ReturnValue As Byte
HIDRetest:
                    DeviceHandle = &HFFFF  'invalid handle initial value
                     
                    ReturnValue = fnGetDeviceHandle(DeviceHandle)
                    Tester.Print ReturnValue; Space(5); ' 1: pass the other refer btnstatus.h
                    Tester.Print "DeviceHandle="; DevicehHandle
                     
                    If ReturnValue <> 1 Then
                        rv0 = UNKNOW       '---> HID mode unknow device mode
                        Call LabelMenu(0, rv0, 1)
                        Tester.Label9.Caption = "HID mode unknow device"
                        fnFreeDeviceHandle (DeviceHandle)
                        GoTo AU6435DLFResult
                    End If
                     
                    '=======================
                    '  key press test, it will return 10 when key up, GPI 6 must do low go hi action
                    '========================

                
                    Do
                        CardResult = DO_WritePort(card, Channel_P1A, &H6E) 'GPI6 : bit 6: pull high
                        Sleep (200)
                        CardResult = DO_WritePort(card, Channel_P1A, &H2E)  ' GPI6 : bit 6: pull low
                        Sleep (1000)
                       
                        ReturnValue = fnInquiryBtnStatus(DeviceHandle)
                        Tester.Print i; Space(5); "Key press value="; ReturnValue
                        i = i + 1
                    Loop While i < 3 And ReturnValue <> 10
                     
                    If (ReturnValue = 12) And ((i = 3) Or (i = 4)) Then
                         
                        GoTo HIDRetest
                    End If
                    
                    If ReturnValue <> 10 Then
                     
                        rv1 = 2
                        Call LabelMenu(1, rv1, rv0)
                        Tester.Label9.Caption = "KeyPress Fail"
                       
                    End If
                              
                End If
                
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print "LBA="; LBA
                 
               
AU6435DLFResult:
                
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

Public Sub AU6435AFF21TestSub()

'insert SD 3.0 Memory Card

Dim TmpLBA As Long
Dim i As Integer
Dim Detect_Counter As Integer
      
Tester.Print "AU6435AF : Begin Test ..."
'==================================================================
'
'  this code come from AU6433DLF20TestSub
'  Add HID test function
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
                    
                AU6371EL_BootTime = 0.1
                     
                    
               If PCI7248InitFinish = 0 Then
                  PCI7248Exist
               End If
               
               '  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
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
                
                If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                End If
                
                'Call MsecDelay(0.2)
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                Call MsecDelay(0.6)    'power on time
                
                ChipString = "vid"
                 
                               
                 '================================================
                'CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
                      
                  
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
                
                Call MsecDelay(0.2)
                
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                
                If CardResult <> 0 Then
                    MsgBox "Read light On fail"
                    End
                End If
                      
                OpenPipe
                rv0 = AU6435Close_OverCurrent(rv0)
                ClosePipe
                      
                If rv0 = 1 Then
                    rv0 = AU6435_CBWTest_New(0, 1, ChipString)
                End If
                
                
                If rv0 = 1 Then
                    rv0 = Read_SD30_Speed_AU6435(0, 0, 64, "4Bits")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
                    End If
                End If
                
                If rv0 = 1 Then
                    rv0 = Read_SD30_Mode_AU6435(0, 0, 64, "SDR")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD3.0 Mode Fail"
                    End If
                End If
                
                
                ClosePipe
                      
                Tester.Print "rv0="; rv0
                     
                If rv0 <> 0 Then
                    If LightOn <> &HF0 Then
                        Tester.Print "Detect Internal Power: "; LightOn
                '        Tester.Print "LightOFF="; LightOFF
                        Tester.Print "V33 or V18 out of range"
                        UsbSpeedTestResult = GPO_FAIL
                        rv0 = 3
                    Else
                        Tester.Print "V18-out Range PASS"
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
                
                If rv0 = 1 Then
                    rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                End If
                                
                If rv1 <> 1 Then
                    LBA = TmpLBA
                    GoTo AU6435DLFResult
                End If
                
                'Next
                LBA = TmpLBA
                      
'=======================================================================================
                        
                    
                     
                Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\128Sector :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                Call MsecDelay(0.1)
                
                rv2 = 0
                Detect_Counter = 0
                
                Do
                    If GetDeviceName_NoReply(ChipString) = "" Then
                        rv2 = 1
                        Tester.Print "NBMD Test PASS!"
                    End If
                    
                    Detect_Counter = Detect_Counter + 1
                    Call MsecDelay(0.05)
                    'Tester.Print Detect_Counter
                Loop Until (Detect_Counter > 10) Or (rv2 = 1)
                
                If rv2 <> 1 Then
                    rv0 = 0
                    Tester.Print "NBMD Test Fail !"
                    GoTo AU6435DLFResult
                End If
                
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
              
                ' No SMC slot
               
                rv2 = rv1   ' to complete the SMC asbolish
               
                
                '===============================================
                '  XD Card test
                '================================================
                ' No XD Slot
                
                rv3 = rv2
                
                '===============================================
                '  MS Card test
                '================================================
                'No MS Slot
                     
                rv4 = rv3
               
                '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  MS Pro Card test
                '================================================
                'No MSpro Slot
                rv5 = rv4
                 
               
AU6435DLFResult:
                
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
Public Sub AU6435BFF20TestSub()

'insert SD 3.0 Memory Card
'2011/3/3 : Add Force & test SD2.0

Dim TmpLBA As Long
Dim i As Integer
Dim Detect_Counter As Integer

Call PowerSet2(1, "3.3", "0.7", 1, "3.3", "0.7", 1)

Tester.Print "AU6435BF : 3.3 Begin Test ..."
'==================================================================
'
'  this code come from AU6433DLF20TestSub
'  Add HID test function
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
                    
                AU6371EL_BootTime = 0.1
                     
                    
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
                'CardResult = DO_WritePort(card, Channel_P1A, &H80)
                
                If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                End If
                
                'Call MsecDelay(0.2)
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                Call MsecDelay(0.3)    'power on time
                
                ChipString = "vid"
                 
                               
                 '================================================
                'CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
                      
                  
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
                
                Call MsecDelay(0.01)
                
                
                
                            
                If rv0 = 1 Then
                    rv0 = AU6435_CBWTest_New(0, 1, ChipString)
                End If
                
                
                If rv0 = 1 Then
                    rv0 = Read_SD30_Speed_AU6435(0, 0, 64, "4Bits")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
                    End If
                End If
                
                If rv0 = 1 Then
                    rv0 = Read_SD30_Mode_AU6435(0, 0, 64, "SDR")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD3.0 Mode Fail"
                    End If
                End If
                
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                    
                If CardResult <> 0 Then
                    MsgBox "Read light On fail"
                    End
                End If
                
                If rv0 <> 0 Then
                    If LightOn <> &HF0 Then
                        Tester.Print "Detect Internal Power: "; LightOn
                '        Tester.Print "LightOFF="; LightOFF
                        Tester.Print "V33 or V18 out of range"
                        UsbSpeedTestResult = GPO_FAIL
                        rv0 = 3
                    Else
                        Tester.Print "V18-out Range PASS"
                    End If
                End If
                
                ClosePipe
                      
                
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      
                TmpLBA = LBA
                LBA = LBA + 199
                            
                ClosePipe
                
                If rv0 = 1 Then
                    rv0 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                End If
                                
                If rv0 <> 1 Then
                    LBA = TmpLBA
                    GoTo AU6435DLFResult
                End If
                
                Tester.Print "rv0="; rv0
                
                Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                LBA = TmpLBA
                      
            '===============================================
            '  SDHC test
            '================================================
                
                Tester.Print "Force SD Card to SDHC Mode (Non-Ultra High Speed)"
                OpenPipe
                rv1 = ReInitial(0)
                Call MsecDelay(0.02)
                rv1 = AU6435ForceSDHC(rv0)
                ClosePipe
                
                If rv1 = 1 Then
                    rv1 = AU6435_CBWTest_New(0, 1, ChipString)
                End If
                
                
                If rv1 = 1 Then
                    rv1 = Read_SD30_Mode_AU6435(0, 0, 64, "Non-UHS")
                    If rv1 <> 1 Then
                        rv1 = 2
                        Tester.Print "SD2.0 Mode Fail"
                    End If
                End If
                
                ClosePipe
                
                Call LabelMenu(1, rv1, rv0)
            
                Tester.Print rv1, " \\SDHC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                 
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                Call MsecDelay(0.1)
                
                If GetDeviceName_NoReply(ChipString) = "" Then
                    rv2 = 1
                    Tester.Print "NBMD Test PASS!"
                End If
                
                If rv2 <> 1 Then
                    rv0 = 0
                    Tester.Print "NBMD Test Fail !"
                    GoTo AU6435DLFResult
                End If
                
                '===============================================
                '  CF Card test
                '================================================
               
                'rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                ' No SMC slot
               
                'rv2 = rv1   ' to complete the SMC asbolish
               
                
                '===============================================
                '  XD Card test
                '================================================
                ' No XD Slot
                
                rv3 = rv2
                
                '===============================================
                '  MS Card test
                '================================================
                'No MS Slot
                     
                rv4 = rv3
               
                '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  MS Pro Card test
                '================================================
                'No MSpro Slot
                rv5 = rv4
                 
               
AU6435DLFResult:
                
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
Public Sub AU6435BFF21TestSub()

'insert SD 3.0 Memory Card
'2011/3/3 : Add Force & test SD2.0

Dim TmpLBA As Long
Dim i As Integer
Dim Detect_Counter As Integer

Tester.Print "AU6435BF : 3.3V Begin Test ..."
'==================================================================
'
'  this code come from AU6433DLF20TestSub
'  Add HID test function
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
                    
                AU6371EL_BootTime = 0.1
                     
                    
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
                'CardResult = DO_WritePort(card, Channel_P1A, &H80)
                
                If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                End If
                
                'Call MsecDelay(0.2)
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                Call MsecDelay(0.3)    'power on time
                
                ChipString = "vid"
                 
                               
                 '================================================
                'CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
                      
                  
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
                
                Call MsecDelay(0.01)
                
                
                
                            
                If rv0 = 1 Then
                    rv0 = AU6435_CBWTest_New(0, 1, ChipString)
                End If
                
                
                If rv0 = 1 Then
                    rv0 = Read_SD30_Speed_AU6435(0, 0, 64, "4Bits")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
                    End If
                End If
                
                If rv0 = 1 Then
                    rv0 = Read_SD30_Mode_AU6435(0, 0, 64, "SDR")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD3.0 Mode Fail"
                    End If
                End If
                
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                    
                If CardResult <> 0 Then
                    MsgBox "Read light On fail"
                    End
                End If
                
                If rv0 <> 0 Then
                    If LightOn <> &HF3 Then
                        Tester.Print "Detect Internal Power: "; LightOn
                '        Tester.Print "LightOFF="; LightOFF
                        Tester.Print "V33 or V18 out of range"
                        UsbSpeedTestResult = GPO_FAIL
                        rv0 = 3
                    Else
                        Tester.Print "V18-out Range PASS"
                    End If
                End If
                
                ClosePipe
                      
                
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      
                TmpLBA = LBA
                LBA = LBA + 199
                            
                ClosePipe
                
                If rv0 = 1 Then
                    rv0 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                End If
                                
                If rv0 <> 1 Then
                    LBA = TmpLBA
                    GoTo AU6435DLFResult
                End If
                
                Tester.Print "rv0="; rv0
                
                Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                LBA = TmpLBA
                      
            '===============================================
            '  SDHC test
            '================================================
                
                Tester.Print "Force SD Card to SDHC Mode (Non-Ultra High Speed)"
                OpenPipe
                rv1 = ReInitial(0)
                Call MsecDelay(0.02)
                rv1 = AU6435ForceSDHC(rv0)
                ClosePipe
                
                If rv1 = 1 Then
                    rv1 = AU6435_CBWTest_New(0, 1, ChipString)
                End If
                
                
                If rv1 = 1 Then
                    rv1 = Read_SD30_Mode_AU6435(0, 0, 64, "Non-UHS")
                    If rv1 <> 1 Then
                        rv1 = 2
                        Tester.Print "SD2.0 Mode Fail"
                    End If
                End If
                
                ClosePipe
                
                Call LabelMenu(1, rv1, rv0)
            
                Tester.Print rv1, " \\SDHC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                 
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                Call MsecDelay(0.1)
                
                If GetDeviceName_NoReply(ChipString) = "" Then
                    rv2 = 1
                    Tester.Print "NBMD Test PASS!"
                End If
                
                If rv2 <> 1 Then
                    rv0 = 0
                    Tester.Print "NBMD Test Fail !"
                    GoTo AU6435DLFResult
                End If
                
                '===============================================
                '  CF Card test
                '================================================
               
                'rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                ' No SMC slot
               
                'rv2 = rv1   ' to complete the SMC asbolish
               
                
                '===============================================
                '  XD Card test
                '================================================
                ' No XD Slot
                
                rv3 = rv2
                
                '===============================================
                '  MS Card test
                '================================================
                'No MS Slot
                     
                rv4 = rv3
               
                '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  MS Pro Card test
                '================================================
                'No MSpro Slot
                rv5 = rv4
                 
               
AU6435DLFResult:
                
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
Public Sub AU6435BFF23TestSub()

'insert SD 3.0 Memory Card
'2011/7/11 : Add Force & test SD2.0

Dim TmpLBA As Long
Dim i As Integer
Dim Detect_Counter As Integer

Tester.Print "AU6435BF : 3.3V Begin Test ..."
'==================================================================
'
'  this code come from AU6433DLF20TestSub
'  Add HID test function
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
                    
                AU6371EL_BootTime = 0.1
                     
                    
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
                'CardResult = DO_WritePort(card, Channel_P1A, &H80)
                
                If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                End If
                
                'Call MsecDelay(0.2)
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                Call MsecDelay(0.3)    'power on time
                
                ChipString = "vid_058f"
                 
                               
                 '================================================
                'CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
                      
                  
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
                
                Call MsecDelay(0.2)
                
                'If rv0 = 1 Then
                '    OpenPipe
                '    rv0 = AU6435Set_Pad_Driving54(1)        'AU6435D51 set driving match customer setting
                '    ClosePipe
                'End If
                
                            
                If rv0 = 1 Then
                    rv0 = CBWTest_New(0, 1, ChipString)
                    If rv0 <> 1 Then
                        Tester.Print "R/W Fail"
                    End If
                End If
                
                
                If rv0 = 1 Then
                    rv0 = Read_SD30_Speed_AU6435_100(0, 0, 64, "4Bits")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
                    End If
                End If
                
                If rv0 = 1 Then
                    rv0 = Read_SD30_Mode_AU6435(0, 0, 64, "SDR")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD3.0 Mode Fail"
                    End If
                End If
                
                Call MsecDelay(0.2)
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                Call MsecDelay(0.2)
                If CardResult <> 0 Then
                    MsgBox "Read light On fail"
                    End
                End If
                
                If rv0 = 1 Then
                    If LightOn <> &HF3 Then
                        Tester.Print "Detect Internal Power: "; LightOn
                '        Tester.Print "LightOFF="; LightOFF
                        Tester.Print "V33 or V18 out of range"
                        UsbSpeedTestResult = GPO_FAIL
                        rv0 = 3
                    Else
                        Tester.Print "V18-out Range PASS"
                    End If
                End If
                
                ClosePipe
                      
                
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      
                TmpLBA = LBA
                LBA = LBA + 199
                            
                ClosePipe
                
                If rv0 = 1 Then
                    rv0 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                    If rv0 <> 1 Then
                        Tester.Print "R/W 64k Fail"
                    End If
                End If
                                
                If rv0 <> 1 Then
                    LBA = TmpLBA
                    GoTo AU6435DLFResult
                End If
                
                Tester.Print "rv0="; rv0
                
                Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                LBA = TmpLBA
                      
            '===============================================
            '  SDHC test
            '================================================
                
                Tester.Print "Force SD Card to SDHC Mode (Non-Ultra High Speed)"
                OpenPipe
                rv1 = ReInitial(0)
                Call MsecDelay(0.02)
                rv1 = AU6435ForceSDHC(rv0)
                ClosePipe
                
                If rv1 = 1 Then
                    rv1 = AU6435_CBWTest_New(0, 1, ChipString)
                End If
                
                
                If rv1 = 1 Then
                    rv1 = Read_SD30_Mode_AU6435(0, 0, 64, "Non-UHS")
                    If rv1 <> 1 Then
                        rv1 = 2
                        Tester.Print "SD2.0 Mode Fail"
                    End If
                End If
                
                ClosePipe
                
                Call LabelMenu(1, rv1, rv0)
            
                Tester.Print rv1, " \\SDHC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                 
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                Call MsecDelay(0.1)
                
                If GetDeviceName_NoReply(ChipString) = "" Then
                    rv2 = 1
                    Tester.Print "NBMD Test PASS!"
                End If
                
                If rv2 <> 1 Then
                    rv0 = 0
                    Tester.Print "NBMD Test Fail !"
                    GoTo AU6435DLFResult
                End If
                
                '===============================================
                '  CF Card test
                '================================================
               
                'rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                ' No SMC slot
               
                'rv2 = rv1   ' to complete the SMC asbolish
               
                
                '===============================================
                '  XD Card test
                '================================================
                ' No XD Slot
                
                rv3 = rv2
                
                '===============================================
                '  MS Card test
                '================================================
                'No MS Slot
                     
                rv4 = rv3
               
                '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  MS Pro Card test
                '================================================
                'No MSpro Slot
                rv5 = rv4
                 
               
AU6435DLFResult:
                
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

Public Sub AU6435CFF22TestSub()

'insert SD 3.0 Memory Card
'2011/7/22 : Add Force & test SD2.0 (copy from AU6433BFF23) set driving


Dim TmpLBA As Long
Dim i As Integer
Dim Detect_Counter As Integer

Tester.Print "AU6435BF : 3.3V Begin Test ..."
'==================================================================
'
'  this code come from AU6433DLF20TestSub
'  Add HID test function
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
                    
                AU6371EL_BootTime = 0.1
                     
                    
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
                Call MsecDelay(0.2)    'power on time
                
                ChipString = "vid_058f"
                 
                 '================================================
                'CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
                      
                  
                'If CardResult <> 0 Then
                '    MsgBox "Read light off fail"
                '    End
                'End If
                   
                 
               
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
                
                Call MsecDelay(0.1)
                
                rv0 = WaitDevOn(ChipString)
                
                Call MsecDelay(0.3)
                
                If rv0 = 1 Then
                    OpenPipe
                    rv0 = AU6435Set_Pad_Driving54(1)        'AU6435D51 set driving match customer setting
                    ClosePipe
                End If
                
                            
                If rv0 = 1 Then
                    ClosePipe
                    rv0 = AU6435_CBWTest_New(0, 1, ChipString)
                    If rv0 <> 1 Then
                        Tester.Print "R/W Fail"
                    End If
                End If
                
                
                If rv0 = 1 Then
                    rv0 = Read_SD30_Speed_AU6435(0, 0, 64, "4Bits")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
                    End If
                End If
                
                If rv0 = 1 Then
                    rv0 = Read_SD30_Mode_AU6435(0, 0, 64, "SDR")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD3.0 Mode Fail"
                    End If
                End If
                
                Call MsecDelay(0.2)
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                Call MsecDelay(0.2)
                If CardResult <> 0 Then
                    MsgBox "Read light On fail"
                    End
                End If
                
                If rv0 = 1 Then
                    If LightOn <> &HE3 Then
                        Tester.Print "Detect Internal Power: "; LightOn
                        Tester.Print "V33 or V18 out of range"
                        UsbSpeedTestResult = GPO_FAIL
                        rv0 = 3
                    Else
                        Tester.Print "V18-out Range PASS"
                    End If
                End If
                
                ClosePipe
                      
                
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      
                TmpLBA = LBA
                LBA = LBA + 199
                            
                ClosePipe
                Call MsecDelay(0.2)
                
                If rv0 = 1 Then
                    rv0 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                    If rv0 <> 1 Then
                        Tester.Print "R/W 64k Fail"
                    End If
                End If
                                
               
                Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                LBA = TmpLBA
                
                If rv0 <> 1 Then
                '    LBA = TmpLBA
                    GoTo AU6435DLFResult
                End If
                      
            '===============================================
            '  SDHC test
            '================================================
                
                Tester.Print "Force SD Card to SDHC Mode (Non-Ultra High Speed)"
                OpenPipe
                rv1 = ReInitial(0)
                Call MsecDelay(0.02)
                rv1 = AU6435ForceSDHC(rv0)
                ClosePipe
                
                If rv1 = 1 Then
                    rv1 = AU6435_CBWTest_New(0, 1, ChipString)
                End If
                
                
                If rv1 = 1 Then
                    rv1 = Read_SD30_Mode_AU6435(0, 0, 64, "Non-UHS")
                    If rv1 <> 1 Then
                        rv1 = 2
                        Tester.Print "SD2.0 Mode Fail"
                    End If
                End If
                
                ClosePipe
                
                Call LabelMenu(1, rv1, rv0)
            
                Tester.Print rv1, " \\SDHC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                 
                rv2 = rv1
                'AU6435GCF no NB-Mode
                'CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                'Call MsecDelay(0.1)
                
                'If GetDeviceName_NoReply(ChipString) = "" Then
                '    rv2 = 1
                '    Tester.Print "NBMD Test PASS!"
                'End If
                
                'If rv2 <> 1 Then
                '    rv0 = 0
                '    Tester.Print "NBMD Test Fail !"
                '    GoTo AU6435DLFResult
                'End If
                
                '===============================================
                '  CF Card test
                '================================================
               
                'rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                ' No SMC slot
               
                'rv2 = rv1   ' to complete the SMC asbolish
               
                
                '===============================================
                '  XD Card test
                '================================================
                ' No XD Slot
                
                rv3 = rv2
                
                '===============================================
                '  MS Card test
                '================================================
                'No MS Slot
                     
                rv4 = rv3
               
                '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  MS Pro Card test
                '================================================
                'No MSpro Slot
                rv5 = rv4
                 
               
AU6435DLFResult:
                
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
Public Sub AU6435BFS11TestSub()

'insert SD 3.0 Memory Card
'2011/3/3 : Add Force & test SD2.0

Dim TmpLBA As Long
Dim i As Integer
Dim Detect_Counter As Integer
Call PowerSet2(1, "5.0", "0.7", 1, "5.0", "0.7", 1)

Tester.Print "AU6435BF : 5V Begin Test ..."
'==================================================================
'
'  this code come from AU6433DLF20TestSub
'  Add HID test function
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
                    
                AU6371EL_BootTime = 0.1
                     
                    
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
                'CardResult = DO_WritePort(card, Channel_P1A, &H80)
                
                If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                End If
                
                'Call MsecDelay(0.2)
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                Call MsecDelay(0.3)    'power on time
                
                ChipString = "vid"
                 
                               
                 '================================================
                'CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
                      
                  
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
                
                Call MsecDelay(0.01)
                
                
                
                            
                If rv0 = 1 Then
                    rv0 = AU6435_CBWTest_New(0, 1, ChipString)
                End If
                
                
                If rv0 = 1 Then
                    rv0 = Read_SD30_Speed_AU6435(0, 0, 64, "4Bits")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
                    End If
                End If
                
                If rv0 = 1 Then
                    rv0 = Read_SD30_Mode_AU6435(0, 0, 64, "SDR")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD3.0 Mode Fail"
                    End If
                End If
                
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                    
                If CardResult <> 0 Then
                    MsgBox "Read light On fail"
                    End
                End If
                
                If rv0 <> 0 Then
                    If LightOn <> &HF0 Then
                        Tester.Print "Detect Internal Power: "; LightOn
                '        Tester.Print "LightOFF="; LightOFF
                        Tester.Print "V33 or V18 out of range"
                        UsbSpeedTestResult = GPO_FAIL
                        rv0 = 3
                    Else
                        Tester.Print "V18-out Range PASS"
                    End If
                End If
                
                ClosePipe
                      
                
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      
                TmpLBA = LBA
                LBA = LBA + 199
                            
                ClosePipe
                
                If rv0 = 1 Then
                    rv0 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                End If
                                
                If rv0 <> 1 Then
                    LBA = TmpLBA
                    GoTo AU6435DLFResult
                End If
                
                Tester.Print "rv0="; rv0
                
                Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                LBA = TmpLBA
                      
            '===============================================
            '  SDHC test
            '================================================
                
                Tester.Print "Force SD Card to SDHC Mode (Non-Ultra High Speed)"
                OpenPipe
                rv1 = ReInitial(0)
                Call MsecDelay(0.02)
                rv1 = AU6435ForceSDHC(rv0)
                ClosePipe
                
                If rv1 = 1 Then
                    rv1 = AU6435_CBWTest_New(0, 1, ChipString)
                End If
                
                
                If rv1 = 1 Then
                    rv1 = Read_SD30_Mode_AU6435(0, 0, 64, "Non-UHS")
                    If rv1 <> 1 Then
                        rv1 = 2
                        Tester.Print "SD2.0 Mode Fail"
                    End If
                End If
                
                ClosePipe
                
                Call LabelMenu(1, rv1, rv0)
            
                Tester.Print rv1, " \\SDHC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                 
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                Call MsecDelay(0.1)
                
                If GetDeviceName_NoReply(ChipString) = "" Then
                    rv2 = 1
                    Tester.Print "NBMD Test PASS!"
                End If
                
                If rv2 <> 1 Then
                    rv0 = 0
                    Tester.Print "NBMD Test Fail !"
                    GoTo AU6435DLFResult
                End If
                
                '===============================================
                '  CF Card test
                '================================================
               
                'rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
                
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                '    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                ' No SMC slot
               
                'rv2 = rv1   ' to complete the SMC asbolish
               
                
                '===============================================
                '  XD Card test
                '================================================
                ' No XD Slot
                
                rv3 = rv2
                
                '===============================================
                '  MS Card test
                '================================================
                'No MS Slot
                     
                rv4 = rv3
               
                '    Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  MS Pro Card test
                '================================================
                'No MSpro Slot
                rv5 = rv4
                 
               
AU6435DLFResult:
                
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
Public Sub AU6435DLF23TestSub()

'insert SD 3.0 Memory Card
'99/12/01 add V18-out detect
Dim TmpLBA As Long
Dim i As Integer
Dim Detect_Counter As Integer
'Call PowerSet2(1, "1.87", "0.05", 1, "1.53", "0.05", 1)
    
Tester.Print "AU6435DL : Begin Test ..."
'==================================================================
'
'  this code come from AU6433DLF20TestSub
'  Add HID test function
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
               
               '  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
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
                
                If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                End If
                
                'Call MsecDelay(0.2)
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                Call MsecDelay(0.2)     'power on time
                
                ChipString = "vid"
                 
               
                
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
                CardResult = DO_WritePort(card, Channel_P1A, &H3E)
                      
                If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect Down Fail"
                    End
                End If
                
                Call MsecDelay(0.3)
                
                rv0 = WaitDevOn(ChipString)
                
                Call MsecDelay(0.4)
                
                OpenPipe
                'rv0 = AU6435Close_OverCurrent(rv0)
                rv0 = AU6435Set_Pad_Driving27(rv0)
                ClosePipe
                
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                
                If CardResult <> 0 Then
                    MsgBox "Read light On fail"
                    End
                End If
                      
                'ClosePipe
                If rv0 = 1 Then
                    rv0 = AU6435_CBWTest_New(0, 1, ChipString)
                End If
                
                
                
                If rv0 = 1 Then
                    rv0 = Read_SD30_Speed_AU6435(0, 0, 64, "4Bits")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
                    End If
                End If
                
                If rv0 = 1 Then
                    rv0 = Read_SD30_Mode_AU6435(0, 0, 64, "SDR")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD3.0 Mode Fail"
                    End If
                End If
                
                
                ClosePipe
                      
                Tester.Print "rv0="; rv0
                     
                If rv0 <> 0 Then
                    If LightOn <> &H13 Or LightOff <> &H93 Then
                        Tester.Print "LightON="; LightOn
                        Tester.Print "LightOFF="; LightOff
                        Tester.Print "GPO_Fail or V18 out of range"
                        UsbSpeedTestResult = GPO_FAIL
                        rv0 = 3
                    Else
                        Tester.Print "V18-out Range PASS"
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
                
                If rv0 = 1 Then
                    rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                End If
                                
                If rv1 <> 1 Then
                    LBA = TmpLBA
                    GoTo AU6435DLFResult
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
               
                'rv3 = 1
                'rv4 = 1
                'rv5 = 1
                'GoTo AU6435tmpResult
                '===============================================
                '  XD Card test
                '================================================
                CardResult = DO_WritePort(card, Channel_P1A, &H36) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                End If
                  
                  
                Call MsecDelay(0.1)
                
                CardResult = DO_WritePort(card, Channel_P1A, &H37)  'XD
                
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
              
                
                
                CardResult = DO_WritePort(card, Channel_P1A, &H37)      'MS + XD
              
                Call MsecDelay(0.1)
               
                CardResult = DO_WritePort(card, Channel_P1A, &H1F)      'MS
               
                 
                Call MsecDelay(0.1)
                
                OpenPipe
                rv5 = ReInitial(0)
                ClosePipe
                
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                
                
                    If rv5 = 1 Then
                        rv5 = Read_MS_Speed_AU6435(0, 0, 64, "4Bits")
                        
                        If rv5 <> 1 Then
                            rv5 = 2
                            Tester.Print "MS bus width Fail"
                        End If
                    End If
                
                Call LabelMenu(31, rv5, rv4)
                
                Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                ClosePipe
                
                If rv5 <> 1 Then
                    GoTo AU6435DLFResult
                End If
                
                CardResult = DO_WritePort(card, Channel_P1A, &H3F)
                Call MsecDelay(0.2)
                
                If GetDeviceName_NoReply(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6435DLFResult
                End If
                                
                'GoTo AU6435tmpResult
                                
                '=================================================================================
                ' HID mode and reader mode ---> compositive device
                If rv5 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &HAF) '  pwr off  for HID mode
                    'result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                    
                    
                    
                    Call MsecDelay(0.2)
          
                    CardResult = DO_WritePort(card, Channel_P1A, &H6E) ' HID mode   'PID_6466
                            
                    Call MsecDelay(0.3)
                    
                    rv5 = WaitDevOn(ChipString)
                    Call MsecDelay(0.2)
                    Detect_Counter = 0
                    
                    LightOn = 0
                    
                    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
          
                    If rv5 = 1 And (LightOn <> &H13 Or LightOff <> &H93) Then
                        UsbSpeedTestResult = GPO_FAIL
                        rv4 = 2
                        Tester.Label9.Caption = "GPO FAIL " & LightOff
                    End If
                
                End If
          
                If rv4 = 1 Then
                ' code begin
                     
                    Tester.Cls
                    Tester.Print "keypress test begin---------------"
                    Dim ReturnValue As Byte
HIDRetest:
                    DeviceHandle = &HFFFF  'invalid handle initial value
                     
                    ReturnValue = fnGetDeviceHandle(DeviceHandle)
                    Tester.Print ReturnValue; Space(5); ' 1: pass the other refer btnstatus.h
                    Tester.Print "DeviceHandle="; DevicehHandle
                     
                    If ReturnValue <> 1 Then
                        rv0 = UNKNOW       '---> HID mode unknow device mode
                        Call LabelMenu(0, rv0, 1)
                        Tester.Label9.Caption = "HID mode unknow device"
                        fnFreeDeviceHandle (DeviceHandle)
                        GoTo AU6435DLFResult
                    End If
                     
                    '=======================
                    '  key press test, it will return 10 when key up, GPI 6 must do low go hi action
                    '========================

                
                    Do
                        CardResult = DO_WritePort(card, Channel_P1A, &H6E) 'GPI6 : bit 6: pull high
                        Sleep (200)
                        CardResult = DO_WritePort(card, Channel_P1A, &H2E)  ' GPI6 : bit 6: pull low
                        Sleep (1000)
                       
                        ReturnValue = fnInquiryBtnStatus(DeviceHandle)
                        Tester.Print i; Space(5); "Key press value="; ReturnValue
                        i = i + 1
                    Loop While i < 3 And ReturnValue <> 10
                     
                    If (ReturnValue = 12) And ((i = 3) Or (i = 4)) Then
                         
                        GoTo HIDRetest
                    End If
                    
                    If ReturnValue <> 10 Then
                     
                        rv1 = 2
                        Call LabelMenu(1, rv1, rv0)
                        Tester.Label9.Caption = "KeyPress Fail"
                       
                    End If
                              
                End If
AU6435tmpResult:
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print "LBA="; LBA
                 
               
AU6435DLFResult:
                
                CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                'Call PowerSet2(1, "0.0", "0.7", 1, "0.0", "0.7", 1)
                
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

Public Sub AU6435DLF24TestSub()

'insert SD 3.0 Memory Card
'99/12/01 add V18-out detect
'100/2/18 Close Pad_Driving & tune pwr-off delay time
'100/3/1 Just revise LightON & LightOFF value (3.3V-Input)

Dim TmpLBA As Long
Dim i As Integer
Dim Detect_Counter As Integer
    
Tester.Print "AU6435DL : Begin Test ..."
'==================================================================
'
'  this code come from AU6433DLF20TestSub
'  Add HID test function
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
               
               '  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
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
                
                If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                End If
                
                'Call MsecDelay(0.2)
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H3F)
                  
                Call MsecDelay(0.3)     'power on time
                
                ChipString = "vid"
                 
                
                
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
                CardResult = DO_WritePort(card, Channel_P1A, &H3E)
                      
                If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect Down Fail"
                    End
                End If
                
                Call MsecDelay(0.3)
                
                rv0 = WaitDevOn(ChipString)
                
                Call MsecDelay(0.01)
                
                If CardResult <> 0 Then
                    MsgBox "Read light On fail"
                    End
                End If
                      
                ClosePipe
                
                If rv0 = 1 Then
                    rv0 = AU6435_CBWTest_New(0, 1, ChipString)
                End If
                
                If rv0 = 1 Then
                    
                    rv0 = Read_SD30_Speed_AU6435(0, 0, 64, "4Bits")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
                    End If
                End If
                
                If rv0 = 1 Then
                    rv0 = Read_SD30_Mode_AU6435(0, 0, 64, "SDR")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD3.0 Mode Fail"
                    End If
                End If
                
                ClosePipe
                      
                Tester.Print "rv0="; rv0
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      
                TmpLBA = LBA
                rv1 = 0
                LBA = LBA + 199
                            
                ClosePipe
                
                If rv0 = 1 Then
                    rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                    ClosePipe
                End If
                                
                If rv1 <> 1 Then
                    LBA = TmpLBA
                    GoTo AU6435DLFResult
                End If
                
                'Next
                LBA = TmpLBA
                      
'=======================================================================================
                        
                    
                     
                Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                Tester.Print rv0, " \\SDXC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  SDHC test
                '================================================
                
                Tester.Print "Force SD Card to SDHC Mode (Non-Ultra High Speed)"
                OpenPipe
                rv1 = ReInitial(0)
                Call MsecDelay(0.02)
                rv1 = AU6435ForceSDHC(rv0)
                ClosePipe
                
                If rv1 = 1 Then
                    rv1 = AU6435_CBWTest_New(0, 1, ChipString)
                End If
                
                
                If rv1 = 1 Then
                    rv1 = Read_SD30_Mode_AU6435(0, 0, 64, "Non-UHS")
                    If rv1 <> 1 Then
                        rv1 = 2
                        Tester.Print "SD2.0 Mode Fail"
                    End If
                End If
                
                ClosePipe
                
                Call LabelMenu(1, rv1, rv0)
            
                Tester.Print rv1, " \\SDHC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                rv2 = rv1   ' to complete the SMC asbolish
               
                'rv3 = 1
                'rv4 = 1
                'rv5 = 1
                'GoTo AU6435tmpResult
                '===============================================
                '  XD Card test
                '================================================
                CardResult = DO_WritePort(card, Channel_P1A, &H36) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                End If
                  
                  
                Call MsecDelay(0.05)
                
                CardResult = DO_WritePort(card, Channel_P1A, &H37)  'XD
                
                Call MsecDelay(0.05)
                 
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
              
                
                
                CardResult = DO_WritePort(card, Channel_P1A, &H37)      'MS + XD
              
                Call MsecDelay(0.05)
               
                CardResult = DO_WritePort(card, Channel_P1A, &H1F)      'MS
               
                OpenPipe
                rv5 = ReInitial(0)
                ClosePipe
                
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                    If rv5 = 1 Then
                        rv5 = Read_MS_Speed_AU6435(0, 0, 64, "4Bits")
                        
                        If rv5 <> 1 Then
                            rv5 = 2
                            Tester.Print "MS bus width Fail"
                        End If
                    End If
                
                Call LabelMenu(31, rv5, rv4)
                
                Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                ClosePipe
                
       
                CardResult = DO_WritePort(card, Channel_P1A, &H3F)
                Call MsecDelay(0.2)
                
                If GetDeviceName_NoReply(ChipString) <> "" Then             'NBMD fail
                    rv0 = 0
                    GoTo AU6435DLFResult
                End If
                
                '=================================================================================
                ' HID mode and reader mode ---> compositive device
                If rv5 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &HAF) '  pwr off  for HID mode
                    'result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                    
                    Call MsecDelay(0.2)
          
                    CardResult = DO_WritePort(card, Channel_P1A, &H6E) ' HID mode   'PID_6466
                            
                    Call MsecDelay(0.3)
                    rv5 = WaitDevOn(ChipString)
                    Call MsecDelay(0.1)
                    Detect_Counter = 0
                    
                    LightOn = 0
                    
                    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
          
                End If
                    
                If (rv5 = 1) And ((LightOn <> &H13) Or (LightOff <> &H93)) Then
                    Tester.Print "LightON="; LightOn
                    Tester.Print "LightOFF="; LightOff
                    Tester.Print "GPO_Fail or V18 out of range"
                    UsbSpeedTestResult = GPO_FAIL
                    rv0 = 3
                    GoTo AU6435DLFResult
                Else
                    Tester.Print "V18-out Range PASS"
                End If
                
                If rv4 = 1 Then
                ' code begin
                     
                    Tester.Cls
                    Tester.Print "keypress test begin---------------"
                    Dim ReturnValue As Byte
HIDRetest:
                    DeviceHandle = &HFFFF  'invalid handle initial value
                     
                    ReturnValue = fnGetDeviceHandle(DeviceHandle)
                    Tester.Print ReturnValue; Space(5); ' 1: pass the other refer btnstatus.h
                    Tester.Print "DeviceHandle="; DevicehHandle
                     
                    If ReturnValue <> 1 Then
                        rv0 = UNKNOW       '---> HID mode unknow device mode
                        Call LabelMenu(0, rv0, 1)
                        Tester.Label9.Caption = "HID mode unknow device"
                        fnFreeDeviceHandle (DeviceHandle)
                        GoTo AU6435DLFResult
                    End If
                     
                    '=======================
                    '  key press test, it will return 10 when key up, GPI 6 must do low go hi action
                    '========================

                
                    Do
                        CardResult = DO_WritePort(card, Channel_P1A, &H6E) 'GPI6 : bit 6: pull high
                        Sleep (200)
                        CardResult = DO_WritePort(card, Channel_P1A, &H2E)  ' GPI6 : bit 6: pull low
                        Sleep (500)
                       
                        ReturnValue = fnInquiryBtnStatus(DeviceHandle)
                        Tester.Print i; Space(5); "Key press value="; ReturnValue
                        i = i + 1
                    Loop While i < 3 And ReturnValue <> 10
                     
                    If (ReturnValue = 12) And ((i = 3) Or (i = 4)) Then
                         
                        GoTo HIDRetest
                    End If
                    
                    If ReturnValue <> 10 Then
                     
                        rv1 = 2
                        Call LabelMenu(1, rv1, rv0)
                        Tester.Label9.Caption = "KeyPress Fail"
                       
                    End If
                              
                End If
AU6435tmpResult:
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print "LBA="; LBA
                 
               
AU6435DLFResult:
                
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

Public Sub AU6435DLF2DTestSub()

'insert SD 3.0 Memory Card
'99/12/01 add V18-out detect
'100/2/18 Close Pad_Driving & tune pwr-off delay time
'100/3/1 Just revise LightON & LightOFF value (3.3V-Input)

Dim TmpLBA As Long
Dim i As Integer
Dim Detect_Counter As Integer
    
Tester.Print "AU6435DL : Begin Test ..."
'==================================================================
'
'  this code come from AU6433DLF20TestSub
'  Add HID test function
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
               
               '  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
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
                
                If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                End If
                
                'Call MsecDelay(0.2)
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H3F)
                  
                Call MsecDelay(0.3)     'power on time
                
                ChipString = "vid"
                 
                
                
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
                CardResult = DO_WritePort(card, Channel_P1A, &H3E)
                      
                If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect Down Fail"
                    End
                End If
                
                Call MsecDelay(0.3)
                
                rv0 = WaitDevOn(ChipString)
                
                Call MsecDelay(0.01)
                
                If CardResult <> 0 Then
                    MsgBox "Read light On fail"
                    End
                End If
                      
                ClosePipe
                
                If rv0 = 1 Then
                    rv0 = AU6435_CBWTest_New(0, 1, ChipString)
                End If
                
                If rv0 = 1 Then
                    
                    rv0 = Read_SD30_Speed_AU6435(0, 0, 64, "4Bits")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
                    End If
                End If
                
                If rv0 = 1 Then
                    rv0 = Read_SD30_Mode_AU6435(0, 0, 64, "SDR")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD3.0 Mode Fail"
                    End If
                End If
                
                ClosePipe
                      
                Tester.Print "rv0="; rv0
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      
                TmpLBA = LBA
                rv1 = 0
                LBA = LBA + 199
                            
                ClosePipe
                
                If rv0 = 1 Then
                    rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                    ClosePipe
                End If
                                
                If rv1 <> 1 Then
                    LBA = TmpLBA
                    GoTo AU6435DLFResult
                End If
                
                'Next
                LBA = TmpLBA
                      
'=======================================================================================
                        
                    
                     
                Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                Tester.Print rv0, " \\SDXC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  SDHC test
                '================================================
                
                Tester.Print "Force SD Card to SDHC Mode (Non-Ultra High Speed)"
                OpenPipe
                rv1 = ReInitial(0)
                Call MsecDelay(0.02)
                rv1 = AU6435ForceSDHC(rv0)
                ClosePipe
                
                If rv1 = 1 Then
                    rv1 = AU6435_CBWTest_New(0, 1, ChipString)
                End If
                
                
                If rv1 = 1 Then
                    rv1 = Read_SD30_Mode_AU6435(0, 0, 64, "Non-UHS")
                    If rv1 <> 1 Then
                        rv1 = 2
                        Tester.Print "SD2.0 Mode Fail"
                    End If
                End If
                
                ClosePipe
                
                Call LabelMenu(1, rv1, rv0)
            
                Tester.Print rv1, " \\SDHC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                rv2 = rv1   ' to complete the SMC asbolish
               
                'rv3 = 1
                'rv4 = 1
                'rv5 = 1
                'GoTo AU6435tmpResult
                '===============================================
                '  XD Card test
                '================================================
                CardResult = DO_WritePort(card, Channel_P1A, &H36) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                End If
                  
                  
                Call MsecDelay(0.05)
                
                CardResult = DO_WritePort(card, Channel_P1A, &H37)  'XD
                
                Call MsecDelay(0.05)
                 
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
              
                
                
                CardResult = DO_WritePort(card, Channel_P1A, &H37)      'MS + XD
              
                Call MsecDelay(0.05)
               
                CardResult = DO_WritePort(card, Channel_P1A, &H1F)      'MS
               
                OpenPipe
                rv5 = ReInitial(0)
                ClosePipe
                
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                    If rv5 = 1 Then
                        rv5 = Read_MS_Speed_AU6435(0, 0, 64, "4Bits")
                        
                        If rv5 <> 1 Then
                            rv5 = 2
                            Tester.Print "MS bus width Fail"
                        End If
                    End If
                
                Call LabelMenu(31, rv5, rv4)
                
                Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                ClosePipe
                
       
                CardResult = DO_WritePort(card, Channel_P1A, &H3F)
                Call MsecDelay(0.2)
                
                If GetDeviceName_NoReply(ChipString) <> "" Then             'NBMD fail
                    rv0 = 0
                    GoTo AU6435DLFResult
                End If
                
                '=================================================================================
                ' HID mode and reader mode ---> compositive device
                If rv5 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &HAF) '  pwr off  for HID mode
                    'result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                    
                    Call MsecDelay(0.2)
          
                    CardResult = DO_WritePort(card, Channel_P1A, &H6E) ' HID mode   'PID_6466
                            
                    Call MsecDelay(0.3)
                    rv5 = WaitDevOn(ChipString)
                    Call MsecDelay(0.1)
                    Detect_Counter = 0
                    
                    LightOn = 0
                    
                    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
          
                End If
                    
                If (rv5 = 1) And ((LightOn <> &H1F) Or (LightOff <> &H9F)) Then
                    Tester.Print "LightON="; LightOn
                    Tester.Print "LightOFF="; LightOff
                    Tester.Print "GPO_Fail or V18 out of range"
                    UsbSpeedTestResult = GPO_FAIL
                    rv0 = 3
                    GoTo AU6435DLFResult
                Else
                    Tester.Print "V18-out Range PASS"
                End If
                
                If rv4 = 1 Then
                ' code begin
                     
                    Tester.Cls
                    Tester.Print "keypress test begin---------------"
                    Dim ReturnValue As Byte
HIDRetest:
                    DeviceHandle = &HFFFF  'invalid handle initial value
                     
                    ReturnValue = fnGetDeviceHandle(DeviceHandle)
                    Tester.Print ReturnValue; Space(5); ' 1: pass the other refer btnstatus.h
                    Tester.Print "DeviceHandle="; DevicehHandle
                     
                    If ReturnValue <> 1 Then
                        rv0 = UNKNOW       '---> HID mode unknow device mode
                        Call LabelMenu(0, rv0, 1)
                        Tester.Label9.Caption = "HID mode unknow device"
                        fnFreeDeviceHandle (DeviceHandle)
                        GoTo AU6435DLFResult
                    End If
                     
                    '=======================
                    '  key press test, it will return 10 when key up, GPI 6 must do low go hi action
                    '========================

                
                    Do
                        CardResult = DO_WritePort(card, Channel_P1A, &H6E) 'GPI6 : bit 6: pull high
                        Sleep (200)
                        CardResult = DO_WritePort(card, Channel_P1A, &H2E)  ' GPI6 : bit 6: pull low
                        Sleep (500)
                       
                        ReturnValue = fnInquiryBtnStatus(DeviceHandle)
                        Tester.Print i; Space(5); "Key press value="; ReturnValue
                        i = i + 1
                    Loop While i < 3 And ReturnValue <> 10
                     
                    If (ReturnValue = 12) And ((i = 3) Or (i = 4)) Then
                         
                        GoTo HIDRetest
                    End If
                    
                    If ReturnValue <> 10 Then
                     
                        rv1 = 2
                        Call LabelMenu(1, rv1, rv0)
                        Tester.Label9.Caption = "KeyPress Fail"
                       
                    End If
                              
                End If
AU6435tmpResult:
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print "LBA="; LBA
                 
               
AU6435DLFResult:
                
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

Public Sub AU6435ELF21TestSub()

'insert SD 3.0 Memory Card
'99/12/01 add V18-out detect
'100/2/18 Close Pad_Driving & tune pwr-off delay time
'100/3/1 Just revise LightON & LightOFF value (3.3V-Input)

Dim TmpLBA As Long
Dim i As Integer
Dim Detect_Counter As Integer

Call PowerSet2(1, "5.0", "0.7", 1, "5.0", "0.7", 1)
    
Tester.Print "AU6435DL : Begin Test ..."
'==================================================================
'
'  this code come from AU6433DLF20TestSub
'  Add HID test function
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
               
               '  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
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
                'Call MsecDelay(0.2)
                
                If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                End If
                
                'Call MsecDelay(0.2)
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H3F)
                  
                Call MsecDelay(0.3)     'power on time
                
                ChipString = "vid"
                 
                
                
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
                CardResult = DO_WritePort(card, Channel_P1A, &H3E)
                      
                If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect Down Fail"
                    End
                End If
                
                Call MsecDelay(0.3)
                
                rv0 = WaitDevOn(ChipString)
                
                Call MsecDelay(0.3)
                
                If CardResult <> 0 Then
                    MsgBox "Read light On fail"
                    End
                End If
                      
                ClosePipe
                
                If rv0 = 1 Then
                    rv0 = AU6435_CBWTest_New(0, 1, ChipString)
                End If
                
                If rv0 = 1 Then
                    Call MsecDelay(0.02)
                    rv0 = Read_SD30_Speed_AU6435(0, 0, 64, "4Bits")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
                    End If
                End If
                
                If rv0 = 1 Then
                    rv0 = Read_SD30_Mode_AU6435(0, 0, 64, "SDR")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD3.0 Mode Fail"
                    End If
                End If
                
                ClosePipe
                      
                Tester.Print "rv0="; rv0
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      
                TmpLBA = LBA
                rv1 = 0
                LBA = LBA + 199
                            
                ClosePipe
                
                If rv0 = 1 Then
                    rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                    ClosePipe
                End If
                                
                If rv1 <> 1 Then
                    LBA = TmpLBA
                    GoTo AU6435DLFResult
                End If
                
                'Next
                LBA = TmpLBA
                      
'=======================================================================================
                        
                    
                     
                Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                Tester.Print rv0, " \\SDXC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  SDHC test
                '================================================
                
                Tester.Print "Force SD Card to SDHC Mode (Non-Ultra High Speed)"
                OpenPipe
                rv1 = ReInitial(0)
                Call MsecDelay(0.02)
                rv1 = AU6435ForceSDHC(rv0)
                ClosePipe
                
                If rv1 = 1 Then
                    rv1 = AU6435_CBWTest_New(0, 1, ChipString)
                End If
                
                
                If rv1 = 1 Then
                    rv1 = Read_SD30_Mode_AU6435(0, 0, 64, "Non-UHS")
                    If rv1 <> 1 Then
                        rv1 = 2
                        Tester.Print "SD2.0 Mode Fail"
                    End If
                End If
                
                ClosePipe
                
                Call LabelMenu(1, rv1, rv0)
            
                Tester.Print rv1, " \\SDHC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                rv2 = rv1   ' to complete the SMC asbolish
               
                'rv3 = 1
                'rv4 = 1
                'rv5 = 1
                'GoTo AU6435tmpResult
                '===============================================
                '  XD Card test
                '================================================
                CardResult = DO_WritePort(card, Channel_P1A, &H36) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                End If
                  
                  
                Call MsecDelay(0.05)
                
                CardResult = DO_WritePort(card, Channel_P1A, &H37)  'XD
                
                Call MsecDelay(0.05)
                 
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
              
                
                
                CardResult = DO_WritePort(card, Channel_P1A, &H37)      'MS + XD
              
                Call MsecDelay(0.05)
               
                CardResult = DO_WritePort(card, Channel_P1A, &H1F)      'MS
               
                OpenPipe
                rv5 = ReInitial(0)
                ClosePipe
                
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                    If rv5 = 1 Then
                        rv5 = Read_MS_Speed_AU6435(0, 0, 64, "4Bits")
                        
                        If rv5 <> 1 Then
                            rv5 = 2
                            Tester.Print "MS bus width Fail"
                        End If
                    End If
                
                Call LabelMenu(31, rv5, rv4)
                
                Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                ClosePipe
                
       
                CardResult = DO_WritePort(card, Channel_P1A, &H3F)
                Call MsecDelay(0.2)
                
                If GetDeviceName_NoReply(ChipString) <> "" Then             'NBMD fail
                    rv0 = 0
                    GoTo AU6435DLFResult
                End If
                
                '=================================================================================
                ' HID mode and reader mode ---> compositive device
                If rv5 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &HAF) '  pwr off  for HID mode
                    'result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                    
                    Call MsecDelay(0.2)
          
                    CardResult = DO_WritePort(card, Channel_P1A, &H6E) ' HID mode   'PID_6466
                            
                    Call MsecDelay(0.3)
                    rv5 = WaitDevOn(ChipString)
                    Call MsecDelay(0.1)
                    Detect_Counter = 0
                    
                    LightOn = 0
                    
                    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
          
                End If
                    
                If (rv5 = 1) And ((LightOn <> &H1F) Or (LightOff <> &H9F)) Then
                    Tester.Print "LightON="; LightOn
                    Tester.Print "LightOFF="; LightOff
                    Tester.Print "GPO_Fail or V18 out of range"
                    UsbSpeedTestResult = GPO_FAIL
                    rv0 = 3
                    GoTo AU6435DLFResult
                Else
                    Tester.Print "V18-out Range PASS"
                End If
                
                If rv4 = 1 Then
                ' code begin
                     
                    Tester.Cls
                    Tester.Print "keypress test begin---------------"
                    Dim ReturnValue As Byte
HIDRetest:
                    DeviceHandle = &HFFFF  'invalid handle initial value
                     
                    ReturnValue = fnGetDeviceHandle(DeviceHandle)
                    Tester.Print ReturnValue; Space(5); ' 1: pass the other refer btnstatus.h
                    Tester.Print "DeviceHandle="; DevicehHandle
                     
                    If ReturnValue <> 1 Then
                        rv0 = UNKNOW       '---> HID mode unknow device mode
                        Call LabelMenu(0, rv0, 1)
                        Tester.Label9.Caption = "HID mode unknow device"
                        fnFreeDeviceHandle (DeviceHandle)
                        GoTo AU6435DLFResult
                    End If
                     
                    '=======================
                    '  key press test, it will return 10 when key up, GPI 6 must do low go hi action
                    '========================

                
                    Do
                        CardResult = DO_WritePort(card, Channel_P1A, &H6E) 'GPI6 : bit 6: pull high
                        Sleep (200)
                        CardResult = DO_WritePort(card, Channel_P1A, &H2E)  ' GPI6 : bit 6: pull low
                        Sleep (500)
                       
                        ReturnValue = fnInquiryBtnStatus(DeviceHandle)
                        Tester.Print i; Space(5); "Key press value="; ReturnValue
                        i = i + 1
                    Loop While i < 3 And ReturnValue <> 10
                     
                    If (ReturnValue = 12) And ((i = 3) Or (i = 4)) Then
                         
                        GoTo HIDRetest
                    End If
                    
                    If ReturnValue <> 10 Then
                     
                        rv1 = 2
                        Call LabelMenu(1, rv1, rv0)
                        Tester.Label9.Caption = "KeyPress Fail"
                       
                    End If
                              
                End If
AU6435tmpResult:
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print "LBA="; LBA
                 
               
AU6435DLFResult:
                
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

Public Sub AU6435ELF22TestSub()

'insert SD 3.0 Memory Card
'99/12/01 add V18-out detect
'100/2/18 Close Pad_Driving & tune pwr-off delay time
'100/3/1 Just revise LightON & LightOFF value (3.3V-Input)
'100/4/1 Skip HID function test

Dim TmpLBA As Long
Dim i As Integer
Dim Detect_Counter As Integer

'Call PowerSet2(1, "5.0", "0.7", 1, "5.0", "0.7", 1)
    
Tester.Print "AU6435DL : Begin Test ..."
'==================================================================
'
'  this code come from AU6433DLF20TestSub
'  Add HID test function
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
               
               '  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
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
                'Call MsecDelay(0.2)
                
                If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                End If
                
                'Call MsecDelay(0.2)
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H3F)
                  
                Call MsecDelay(0.3)     'power on time
                
                ChipString = "vid"
                 
                
                
                 '================================================
                CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                Call MsecDelay(0.02)
                  
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
                CardResult = DO_WritePort(card, Channel_P1A, &H3E)
                      
                If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect Down Fail"
                    End
                End If
                
                Call MsecDelay(0.3)
                
                rv0 = WaitDevOn(ChipString)
                
                Call MsecDelay(0.3)
                
                If CardResult <> 0 Then
                    MsgBox "Read light On fail"
                    End
                End If
                
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                Call MsecDelay(0.02)
                
                If (LightOn <> &H1F) Or (LightOff <> &H9F) Then
                    Tester.Print "LightON="; LightOn
                    Tester.Print "LightOFF="; LightOff
                    Tester.Print "GPO_Fail or V18 out of range"
                    UsbSpeedTestResult = GPO_FAIL
                    rv0 = 3
                    GoTo AU6435DLFResult
                Else
                    Tester.Print "V18-out Range PASS"
                End If
                
                Tester.Print "LBA="; LBA
                 
                ClosePipe
                
                If rv0 = 1 Then
                    rv0 = AU6435_CBWTest_New(0, 1, ChipString)
                End If
                
                If rv0 = 1 Then
                    Call MsecDelay(0.02)
                    rv0 = Read_SD30_Speed_AU6435(0, 0, 64, "4Bits")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
                    End If
                End If
                
                If rv0 = 1 Then
                    rv0 = Read_SD30_Mode_AU6435(0, 0, 64, "SDR")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD3.0 Mode Fail"
                    End If
                End If
                
                ClosePipe
                      
                Tester.Print "rv0="; rv0
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      
                TmpLBA = LBA
                rv1 = 0
                LBA = LBA + 199
                            
                ClosePipe
                
                If rv0 = 1 Then
                    rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                    ClosePipe
                End If
                                
                If rv1 <> 1 Then
                    LBA = TmpLBA
                    GoTo AU6435DLFResult
                End If
                
                'Next
                LBA = TmpLBA
                      
'=======================================================================================
                        
                Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                Tester.Print rv0, " \\SDXC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  SDHC test
                '================================================
                
                Tester.Print "Force SD Card to SDHC Mode (Non-Ultra High Speed)"
                OpenPipe
                rv1 = ReInitial(0)
                Call MsecDelay(0.02)
                rv1 = AU6435ForceSDHC(rv0)
                ClosePipe
                
                If rv1 = 1 Then
                    rv1 = AU6435_CBWTest_New(0, 1, ChipString)
                End If
                
                
                If rv1 = 1 Then
                    rv1 = Read_SD30_Mode_AU6435(0, 0, 64, "Non-UHS")
                    If rv1 <> 1 Then
                        rv1 = 2
                        Tester.Print "SD2.0 Mode Fail"
                    End If
                End If
                
                ClosePipe
                
                Call LabelMenu(1, rv1, rv0)
            
                Tester.Print rv1, " \\SDHC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                rv2 = rv1   ' to complete the SMC asbolish
               
                'rv3 = 1
                'rv4 = 1
                'rv5 = 1
                'GoTo AU6435tmpResult
                '===============================================
                '  XD Card test
                '================================================
                CardResult = DO_WritePort(card, Channel_P1A, &H36) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                End If
                  
                  
                Call MsecDelay(0.05)
                
                CardResult = DO_WritePort(card, Channel_P1A, &H37)  'XD
                
                Call MsecDelay(0.05)
                 
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
              
                CardResult = DO_WritePort(card, Channel_P1A, &H37)      'MS + XD
              
                Call MsecDelay(0.05)
               
                CardResult = DO_WritePort(card, Channel_P1A, &H1F)      'MS
               
                OpenPipe
                rv5 = ReInitial(0)
                ClosePipe
                
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                    If rv5 = 1 Then
                        rv5 = Read_MS_Speed_AU6435(0, 0, 64, "4Bits")
                        
                        If rv5 <> 1 Then
                            rv5 = 2
                            Tester.Print "MS bus width Fail"
                        End If
                    End If
                
                Call LabelMenu(31, rv5, rv4)
                
                Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                ClosePipe
                
       
                CardResult = DO_WritePort(card, Channel_P1A, &H3F)
                Call MsecDelay(0.2)
                
                If GetDeviceName_NoReply(ChipString) <> "" Then             'NBMD fail
                    rv0 = 0
                End If
                
               
AU6435DLFResult:
                
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

Public Sub AU6435ELF23TestSub()

'insert SD 3.0 Memory Card
'99/12/01 add V18-out detect
'100/2/18 Close Pad_Driving & tune pwr-off delay time
'100/3/1 Just revise LightON & LightOFF value (3.3V-Input)
'100/4/1 Skip HID function test
'100/6/3 AU6435D51 set driving, match customer setting

Dim TmpLBA As Long
Dim i As Integer
Dim Detect_Counter As Integer

'Call PowerSet2(1, "5.0", "0.7", 1, "5.0", "0.7", 1)
    
Tester.Print "AU6435EL : Begin Test ..."
'==================================================================
'
'  this code come from AU6433DLF20TestSub
'  Add HID test function
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
               
               '  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
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
                'Call MsecDelay(0.2)
                
                If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                End If
                
                'Call MsecDelay(0.2)
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H3F)
                  
                Call MsecDelay(0.2)     'power on time
                
                ChipString = "vid"
                 
                
                
                 '================================================
                CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                Call MsecDelay(0.02)
                  
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
                CardResult = DO_WritePort(card, Channel_P1A, &H3E)
                      
                If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect Down Fail"
                    End
                End If
                
                Call MsecDelay(0.2)
                
                rv0 = WaitDevOn(ChipString)
                
                Call MsecDelay(0.2)
                
                If CardResult <> 0 Then
                    MsgBox "Read light On fail"
                    End
                End If
                
                If rv0 = 1 Then
                    OpenPipe
                    rv0 = AU6435Set_Pad_Driving54(1)        'AU6435D51 set driving match customer setting
                    ClosePipe
                End If
                
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                Call MsecDelay(0.02)
                
                If (LightOn <> &H1F) Or (LightOff <> &H9F) Then
                    Tester.Print "LightON="; LightOn
                    Tester.Print "LightOFF="; LightOff
                    Tester.Print "GPO_Fail or V18 out of range"
                    UsbSpeedTestResult = GPO_FAIL
                    rv0 = 3
                    GoTo AU6435DLFResult
                Else
                    Tester.Print "V18-out Range PASS"
                End If
                
                Tester.Print "LBA="; LBA
                 
                ClosePipe
                
                If rv0 = 1 Then
                    rv0 = AU6435_CBWTest_New(0, 1, ChipString)
                End If
                
                If rv0 = 1 Then
                    Call MsecDelay(0.02)
                    rv0 = Read_SD30_Speed_AU6435(0, 0, 64, "4Bits")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
                    End If
                End If
                
                If rv0 = 1 Then
                    rv0 = Read_SD30_Mode_AU6435(0, 0, 64, "SDR")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD3.0 Mode Fail"
                    End If
                End If
                
                ClosePipe
                      
                Tester.Print "rv0="; rv0
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      
                TmpLBA = LBA
                rv1 = 0
                LBA = LBA + 199
                            
                ClosePipe
                
                If rv0 = 1 Then
                    rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                    ClosePipe
                End If
                                
                If rv1 <> 1 Then
                    LBA = TmpLBA
                    GoTo AU6435DLFResult
                End If
                
                'Next
                LBA = TmpLBA
                      
'=======================================================================================
                        
                Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                Tester.Print rv0, " \\SDXC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  SDHC test
                '================================================
                
                Tester.Print "Force SD Card to SDHC Mode (Non-Ultra High Speed)"
                OpenPipe
                rv1 = ReInitial(0)
                Call MsecDelay(0.02)
                rv1 = AU6435ForceSDHC(rv0)
                ClosePipe
                
                If rv1 = 1 Then
                    rv1 = AU6435_CBWTest_New(0, 1, ChipString)
                End If
                
                
                If rv1 = 1 Then
                    rv1 = Read_SD30_Mode_AU6435(0, 0, 64, "Non-UHS")
                    If rv1 <> 1 Then
                        rv1 = 2
                        Tester.Print "SD2.0 Mode Fail"
                    End If
                End If
                
                ClosePipe
                
                Call LabelMenu(1, rv1, rv0)
            
                Tester.Print rv1, " \\SDHC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                rv2 = rv1   ' to complete the SMC asbolish
               
                '===============================================
                '  XD Card test
                '================================================
                CardResult = DO_WritePort(card, Channel_P1A, &H36) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                End If
                  
                  
                Call MsecDelay(0.05)
                
                CardResult = DO_WritePort(card, Channel_P1A, &H37)  'XD
                
                Call MsecDelay(0.05)
                 
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
              
                CardResult = DO_WritePort(card, Channel_P1A, &H37)      'MS + XD
              
                Call MsecDelay(0.05)
               
                CardResult = DO_WritePort(card, Channel_P1A, &H1F)      'MS
               
                OpenPipe
                rv5 = ReInitial(0)
                ClosePipe
                
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                    If rv5 = 1 Then
                        rv5 = Read_MS_Speed_AU6435(0, 0, 64, "4Bits")
                        
                        If rv5 <> 1 Then
                            rv5 = 2
                            Tester.Print "MS bus width Fail"
                        End If
                    End If
                
                Call LabelMenu(31, rv5, rv4)
                
                Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                ClosePipe
                
       
                CardResult = DO_WritePort(card, Channel_P1A, &H3F)
                Call MsecDelay(0.2)
                
                If GetDeviceName_NoReply(ChipString) <> "" Then             'NBMD fail
                    rv0 = 0
                End If
                
               
AU6435DLFResult:
                
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



Public Sub AU6435ELF33TestSub()

'insert SD 3.0 Memory Card
'99/12/01 add V18-out detect
'100/2/18 Close Pad_Driving & tune pwr-off delay time
'100/3/1 Just revise LightON & LightOFF value (3.3V-Input)
'100/4/1 Skip HID function test
'100/6/3 AU6435D51 set driving, match customer setting

Dim TmpLBA As Long
Dim i As Integer
Dim DetectCount As Integer

If PCI7248InitFinish = 0 Then
    PCI7248ExistAU6254
    Call SetTimer_1ms
End If

OS_Result = 0
rv0 = 0
DetectCount = 0

'Call PowerSet2(0, "0.0", "0.5", 1, "0.0", "0.5", 1)
CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
Call MsecDelay(0.02)
CardResult = DO_WritePort(card, Channel_P1C, &H0)
MsecDelay (0.2)

OpenShortTest_Result_AU6435ELF33

If OS_Result <> 1 Then
    rv0 = 0                 'OS Fail
    GoTo AU6435DLFResult
End If

'Call PowerSet2(0, "3.3", "0.7", 1, "3.3", "0.7", 1)
CardResult = DO_WritePort(card, Channel_P1C, &HFF)
MsecDelay (0.02)

    
Tester.Print "AU6435EL : Begin Test ..."
'==================================================================
'
'  this code come from AU6433DLF20TestSub
'  Add HID test function
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
               
               '  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
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
                'Call MsecDelay(0.2)
                
                If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                End If
                
                'Call MsecDelay(0.2)
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H3F)
                  
                Call MsecDelay(0.2)     'power on time
                
                ChipString = "vid"
                 
                
                
                 '================================================
                CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                Call MsecDelay(0.02)
                  
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
                CardResult = DO_WritePort(card, Channel_P1A, &H3E)
                      
                If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect Down Fail"
                    End
                End If
                
                Call MsecDelay(0.1)
                
                rv0 = WaitDevOn(ChipString)
                
                Call MsecDelay(0.2)
                
                If CardResult <> 0 Then
                    MsgBox "Read light On fail"
                    End
                End If
                
                If rv0 = 1 Then
                    OpenPipe
                    rv0 = AU6435Set_Pad_Driving54(1)        'AU6435D51 set driving match customer setting
                    ClosePipe
                End If
                
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                Call MsecDelay(0.02)
                
                If (LightOn <> &H73) Or (LightOff <> &HF3) Then
                    Tester.Print "LightON="; LightOn
                    Tester.Print "LightOFF="; LightOff
                    Tester.Print "GPO_Fail or V18 out of range"
                    UsbSpeedTestResult = GPO_FAIL
                    rv0 = 3
                    GoTo AU6435DLFResult
                Else
                    Tester.Print "V18-out Range PASS"
                End If
                
                Tester.Print "LBA="; LBA
                 
                ClosePipe
                
                If rv0 = 1 Then
                    rv0 = AU6435_CBWTest_New(0, 1, ChipString)
                End If
                
                If rv0 = 1 Then
                    Call MsecDelay(0.02)
                    rv0 = Read_SD30_Speed_AU6435(0, 0, 64, "4Bits")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
                    End If
                End If
                
                If rv0 = 1 Then
                    rv0 = Read_SD30_Mode_AU6435(0, 0, 64, "SDR")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD3.0 Mode Fail"
                    End If
                End If
                
                ClosePipe
                      
                Tester.Print "rv0="; rv0
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      
                TmpLBA = LBA
                rv1 = 0
                LBA = LBA + 199
                            
                ClosePipe
                
                If rv0 = 1 Then
                    rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                    ClosePipe
                End If
                                
                If rv1 <> 1 Then
                    LBA = TmpLBA
                    GoTo AU6435DLFResult
                End If
                
                'Next
                LBA = TmpLBA
                      
'=======================================================================================
                        
                Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                Tester.Print rv0, " \\SDXC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  SDHC test
                '================================================
                
                Tester.Print "Force SD Card to SDHC Mode (Non-Ultra High Speed)"
                OpenPipe
                rv1 = ReInitial(0)
                Call MsecDelay(0.02)
                rv1 = AU6435ForceSDHC(rv0)
                ClosePipe
                
                If rv1 = 1 Then
                    rv1 = AU6435_CBWTest_New(0, 1, ChipString)
                End If
                
                
                If rv1 = 1 Then
                    rv1 = Read_SD30_Mode_AU6435(0, 0, 64, "Non-UHS")
                    If rv1 <> 1 Then
                        rv1 = 2
                        Tester.Print "SD2.0 Mode Fail"
                    End If
                End If
                
                ClosePipe
                
                Call LabelMenu(1, rv1, rv0)
            
                Tester.Print rv1, " \\SDHC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                rv2 = rv1   ' to complete the SMC asbolish
               
                '===============================================
                '  XD Card test
                '================================================
                CardResult = DO_WritePort(card, Channel_P1A, &H36) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                End If
                  
                  
                Call MsecDelay(0.05)
                
                CardResult = DO_WritePort(card, Channel_P1A, &H37)  'XD
                
                Call MsecDelay(0.05)
                 
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
              
                CardResult = DO_WritePort(card, Channel_P1A, &H37)      'MS + XD
              
                Call MsecDelay(0.05)
               
                CardResult = DO_WritePort(card, Channel_P1A, &H1F)      'MS
               
                OpenPipe
                rv5 = ReInitial(0)
                ClosePipe
                
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                    If rv5 = 1 Then
                        rv5 = Read_MS_Speed_AU6435(0, 0, 64, "4Bits")
                        
                        If rv5 <> 1 Then
                            rv5 = 2
                            Tester.Print "MS bus width Fail"
                        End If
                    End If
                
                Call LabelMenu(31, rv5, rv4)
                
                Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                ClosePipe
                
       
                CardResult = DO_WritePort(card, Channel_P1A, &H3F)
                Call MsecDelay(0.2)
                
                If GetDeviceName_NoReply(ChipString) <> "" Then             'NBMD fail
                    rv0 = 0
                End If
                
               
AU6435DLFResult:
                
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

Public Sub AU6435ELF34TestSub()

'insert SD 3.0 Memory Card
'99/12/01 add V18-out detect
'100/2/18 Close Pad_Driving & tune pwr-off delay time
'100/3/1 Just revise LightON & LightOFF value (3.3V-Input)
'100/4/1 Skip HID function test
'100/6/3 AU6435D51 set driving, match customer setting
'100/8/4 Just modify Pin3 OS become open state, Pin2(Ext48in) skip
'101/5/7 Using S/B: "AU6435-DL 48LQ(OS FT) SOCKET V1.9"

Dim TmpLBA As Long
Dim i As Integer
Dim DetectCount As Integer

If PCI7248InitFinish = 0 Then
    PCI7248ExistAU6254
    Call SetTimer_1ms
End If

OS_Result = 0
rv0 = 0
DetectCount = 0

'Call PowerSet2(0, "0.0", "0.5", 1, "0.0", "0.5", 1)
CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
Call MsecDelay(0.02)
CardResult = DO_WritePort(card, Channel_P1C, &H0)
MsecDelay (0.3)

OpenShortTest_Result_AU6435ELF34

If OS_Result <> 1 Then
    rv0 = 0                 'OS Fail
    GoTo AU6435DLFResult
End If

'Call PowerSet2(0, "3.3", "0.7", 1, "3.3", "0.7", 1)
CardResult = DO_WritePort(card, Channel_P1C, &HFF)
MsecDelay (0.02)

    
Tester.Print "AU6435EL : Begin Test ..."
'==================================================================
'
'  this code come from AU6433DLF20TestSub
'  Add HID test function
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
               
               '  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
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
                'Call MsecDelay(0.2)
                
                If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                End If
                
                'Call MsecDelay(0.2)
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H3F)
                  
                Call MsecDelay(0.2)     'power on time
                
                ChipString = "vid"
                 
                
                
                 '================================================
                CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                Call MsecDelay(0.02)
                  
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
                CardResult = DO_WritePort(card, Channel_P1A, &H3E)
                      
                If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect Down Fail"
                    End
                End If
                
                Call MsecDelay(0.1)
                
                rv0 = WaitDevOn(ChipString)
                
                Call MsecDelay(0.2)
                
                If CardResult <> 0 Then
                    MsgBox "Read light On fail"
                    End
                End If
                
                If rv0 = 1 Then
                    OpenPipe
                    rv0 = AU6435Set_Pad_Driving54(1)        'AU6435D51 set driving match customer setting
                    ClosePipe
                Else
                    GoTo AU6435DLFResult
                End If
                
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                Call MsecDelay(0.02)
                
                If (LightOn <> &H73) Or (LightOff <> &HF3) Then
                    Tester.Print "LightON="; LightOn
                    Tester.Print "LightOFF="; LightOff
                    Tester.Print "GPO_Fail or V18 out of range"
                    UsbSpeedTestResult = GPO_FAIL
                    rv0 = 3
                    GoTo AU6435DLFResult
                Else
                    Tester.Print "V18-out Range PASS"
                End If
                
                Tester.Print "LBA="; LBA
                 
                ClosePipe
                
                If rv0 = 1 Then
                    rv0 = AU6435_CBWTest_New(0, 1, ChipString)
                End If
                
                If rv0 = 1 Then
                    Call MsecDelay(0.02)
                    rv0 = Read_SD30_Speed_AU6435(0, 0, 64, "4Bits")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
                    End If
                End If
                
                If rv0 = 1 Then
                    rv0 = Read_SD30_Mode_AU6435(0, 0, 64, "SDR")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD3.0 Mode Fail"
                    End If
                End If
                
                Tester.Print "rv0="; rv0
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      
                TmpLBA = LBA
                rv1 = 0
                LBA = LBA + 199
                            
                If rv0 = 1 Then
                    rv0 = CBWTest_New_128_Sector_PipeReady(0, rv0)  ' write
                    ClosePipe
                End If
                                
                ClosePipe
                
                If rv0 <> 1 Then
                    LBA = TmpLBA
                    GoTo AU6435DLFResult
                End If
                
                'Next
                LBA = TmpLBA
                      
'=======================================================================================
                        
                Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                Tester.Print rv0, " \\SDXC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  SDHC test
                '================================================
                
                Tester.Print "Force SD Card to SDHC Mode (Non-Ultra High Speed)"
                OpenPipe
                rv1 = ReInitial(0)
                Call MsecDelay(0.02)
                rv1 = AU6435ForceSDHC(rv0)
                ClosePipe
                
                If rv1 = 1 Then
                    rv1 = AU6435_CBWTest_New(0, 1, ChipString)
                End If
                
                
                If rv1 = 1 Then
                    rv1 = Read_SD30_Mode_AU6435(0, 0, 64, "Non-UHS")
                    If rv1 <> 1 Then
                        rv1 = 2
                        Tester.Print "SD2.0 Mode Fail"
                    End If
                End If
                
                ClosePipe
                
                Call LabelMenu(1, rv1, rv0)
            
                Tester.Print rv1, " \\SDHC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                rv2 = rv1   ' to complete the SMC asbolish
               
                '===============================================
                '  XD Card test
                '================================================
                CardResult = DO_WritePort(card, Channel_P1A, &H36) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                End If
                  
                  
                Call MsecDelay(0.02)
                
                CardResult = DO_WritePort(card, Channel_P1A, &H37)  'XD
                
                Call MsecDelay(0.2)
                 
                OpenPipe
                rv3 = ReInitial(0)
                ClosePipe
                
                'Call MsecDelay(0.2)
                
                rv3 = AU6435_CBWTest_New(0, rv2, ChipString)
                
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
              
                CardResult = DO_WritePort(card, Channel_P1A, &H37)      'MS + XD
              
                Call MsecDelay(0.05)
               
                CardResult = DO_WritePort(card, Channel_P1A, &H1F)      'MS
               
                OpenPipe
                rv5 = ReInitial(0)
                ClosePipe
                
                rv5 = AU6435_CBWTest_New(0, rv4, ChipString)
                
                    If rv5 = 1 Then
                        rv5 = Read_MS_Speed_AU6435(0, 0, 64, "4Bits")
                        
                        If rv5 <> 1 Then
                            rv5 = 2
                            Tester.Print "MS bus width Fail"
                        End If
                    End If
                
                Call LabelMenu(31, rv5, rv4)
                
                Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                ClosePipe
                
       
                CardResult = DO_WritePort(card, Channel_P1A, &H3F)
                Call MsecDelay(0.2)
                
                If GetDeviceName_NoReply(ChipString) <> "" Then             'NBMD fail
                    rv0 = 0
                End If
                
               
AU6435DLFResult:
                
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


Public Sub AU6435ELF24TestSub()

'insert SD 3.0 Memory Card
'99/12/01 add V18-out detect
'100/2/18 Close Pad_Driving & tune pwr-off delay time
'100/3/1 Just revise LightON & LightOFF value (3.3V-Input)
'100/4/1 Skip HID function test
'100/6/3 AU6435D51 set driving, match customer setting
'100/8/4 Just modify Pin3 OS become open state, Pin2(Ext48in) skip
'2011/11/10: Purpose to solve V1.7 S/B Pin2(Ext48In) test coverage, this version is FT2 & using 48Mhz clock source

Dim TmpLBA As Long
Dim i As Integer
Dim DetectCount As Integer

If PCI7248InitFinish = 0 Then
    PCI7248ExistAU6254
    Call SetTimer_1ms
End If

OS_Result = 0
rv0 = 0
DetectCount = 0

'Call PowerSet2(0, "0.0", "0.5", 1, "0.0", "0.5", 1)
'CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
'Call MsecDelay(0.02)
'CardResult = DO_WritePort(card, Channel_P1C, &H0)
'MsecDelay (0.3)

'OpenShortTest_Result_AU6435ELF34

'If OS_Result <> 1 Then
'    rv0 = 0                 'OS Fail
'    GoTo AU6435DLFResult
'End If

'Call PowerSet2(0, "3.3", "0.7", 1, "3.3", "0.7", 1)
CardResult = DO_WritePort(card, Channel_P1C, &HFF)
MsecDelay (0.02)

    
Tester.Print "AU6435EL : Begin Test ..."
'==================================================================
'
'  this code come from AU6433DLF20TestSub
'  Add HID test function
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
               
               '  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
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
                'Call MsecDelay(0.2)
                
                If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                End If
                
                'Call MsecDelay(0.2)
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H3F)
                  
                Call MsecDelay(0.2)     'power on time
                
                ChipString = "vid"
                 
                
                
                 '================================================
                CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                Call MsecDelay(0.02)
                  
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
                CardResult = DO_WritePort(card, Channel_P1A, &H3E)
                      
                If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect Down Fail"
                    End
                End If
                
                Call MsecDelay(0.1)
                
                rv0 = WaitDevOn(ChipString)
                
                Call MsecDelay(0.2)
                
                If CardResult <> 0 Then
                    MsgBox "Read light On fail"
                    End
                End If
                
                If rv0 = 1 Then
                    OpenPipe
                    rv0 = AU6435Set_Pad_Driving54(1)        'AU6435D51 set driving match customer setting
                    ClosePipe
                Else
                    GoTo AU6435DLFResult
                End If
                
                
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                Call MsecDelay(0.02)
                
                If (LightOn <> &H73) Or (LightOff <> &HF3) Then
                    Tester.Print "LightON="; LightOn
                    Tester.Print "LightOFF="; LightOff
                    Tester.Print "GPO_Fail or V18 out of range"
                    UsbSpeedTestResult = GPO_FAIL
                    rv0 = 3
                    GoTo AU6435DLFResult
                Else
                    Tester.Print "V18-out Range PASS"
                End If
                
                Tester.Print "LBA="; LBA
                 
                ClosePipe
                
                If rv0 = 1 Then
                    rv0 = AU6435_CBWTest_New(0, 1, ChipString)
                End If
                
                If rv0 = 1 Then
                    Call MsecDelay(0.02)
                    rv0 = Read_SD30_Speed_AU6435(0, 0, 64, "4Bits")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
                    End If
                End If
                
                If rv0 = 1 Then
                    rv0 = Read_SD30_Mode_AU6435(0, 0, 64, "SDR")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD3.0 Mode Fail"
                    End If
                End If
                
                'Tester.Print "rv0="; rv0
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      
                TmpLBA = LBA
                rv1 = 0
                LBA = LBA + 199
                            
                If rv0 = 1 Then
                    rv0 = CBWTest_New_128_Sector_PipeReady(0, rv0)  ' write
                    ClosePipe
                End If
                                
                ClosePipe
                
                'If rv0 <> 1 Then
                '    LBA = TmpLBA
                '    GoTo AU6435DLFResult
                'End If
                
                'Next
                LBA = TmpLBA
                      
'=======================================================================================
                        
                Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                Tester.Print rv0, " \\SDXC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  SDHC test
                '================================================
                
                Tester.Print "Force SD Card to SDHC Mode (Non-Ultra High Speed)"
                OpenPipe
                rv1 = ReInitial(0)
                Call MsecDelay(0.02)
                rv1 = AU6435ForceSDHC(rv0)
                ClosePipe
                
                If rv1 = 1 Then
                    rv1 = AU6435_CBWTest_New(0, 1, ChipString)
                End If
                
                
                If rv1 = 1 Then
                    rv1 = Read_SD30_Mode_AU6435(0, 0, 64, "Non-UHS")
                    If rv1 <> 1 Then
                        rv1 = 2
                        Tester.Print "SD2.0 Mode Fail"
                    End If
                End If
                
                ClosePipe
                
                Call LabelMenu(1, rv1, rv0)
            
                Tester.Print rv1, " \\SDHC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                rv2 = rv1   ' to complete the SMC asbolish
               
                '===============================================
                '  XD Card test
                '================================================
                CardResult = DO_WritePort(card, Channel_P1A, &H36) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                End If
                  
                  
                Call MsecDelay(0.05)
                
                CardResult = DO_WritePort(card, Channel_P1A, &H37)  'XD
                
                Call MsecDelay(0.05)
                 
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
              
                CardResult = DO_WritePort(card, Channel_P1A, &H17)      'MS + XD
              
                Call MsecDelay(0.05)
               
                CardResult = DO_WritePort(card, Channel_P1A, &H1F)      'MS
               
                OpenPipe
                rv5 = ReInitial(0)
                ClosePipe
                
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                    If rv5 = 1 Then
                        rv5 = Read_MS_Speed_AU6435(0, 0, 64, "4Bits")
                        
                        If rv5 <> 1 Then
                            rv5 = 2
                            Tester.Print "MS bus width Fail"
                        End If
                    End If
                
                Call LabelMenu(31, rv5, rv4)
                
                Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                ClosePipe
                
       
                CardResult = DO_WritePort(card, Channel_P1A, &H3F)
                Call MsecDelay(0.2)
                
                If GetDeviceName_NoReply(ChipString) <> "" Then             'NBMD fail
                    rv0 = 0
                End If
                
               
AU6435DLFResult:
                
                CardResult = DO_WritePort(card, Channel_P1A, &H3F)   ' Close power
                
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

Public Sub AU6435ELF25TestSub()

'insert SD 3.0 Memory Card
'99/12/01 add V18-out detect
'100/2/18 Close Pad_Driving & tune pwr-off delay time
'100/3/1 Just revise LightON & LightOFF value (3.3V-Input)
'100/4/1 Skip HID function test
'100/6/3 AU6435D51 set driving, match customer setting
'100/8/4 Just modify Pin3 OS become open state, Pin2(Ext48in) skip
'2011/11/10: Purpose to solve V1.7 S/B Pin2(Ext48In) test coverage, this version is FT2 & using 48Mhz clock source
'2012/07/25: for Quanta RMA CSMC IC on smooth power curve will unknow issue

Dim TmpLBA As Long
Dim i As Integer
Dim DetectCount As Integer

If PCI7248InitFinish = 0 Then
    PCI7248ExistAU6254
    Call SetTimer_1ms
End If

OS_Result = 0
rv0 = 0
DetectCount = 0

'CardResult = DO_WritePort(card, Channel_P1C, &HFF)
'MsecDelay (0.02)

    
Tester.Print "AU6435EL : Begin Test ..."
'==================================================================
'
'  this code come from AU6433DLF20TestSub
'  Add HID test function
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
     
    
If PCI7248InitFinish_Sync = 0 Then
  PCI7248Exist_P1C_Sync
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
'CardResult = DO_WritePort(card, Channel_P1A, &H80)

CardResult = DO_WritePort(card, Channel_P1A, &H3F)
Call MsecDelay(0.2)

If CardResult <> 0 Then
    MsgBox "Power off fail"
    End
End If

'Call MsecDelay(0.2)     'power on time

ChipString = "vid"

 '================================================
CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
Call MsecDelay(0.02)
  
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
CardResult = DO_WritePort(card, Channel_P1A, &H3E)
      
If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If

Call PowerSet2(0, "5.0", "0.7", 1, "5.0", "0.7", 1)
Call MsecDelay(0.2)

SetSiteStatus (RunHV)
Call WaitAnotherSiteDone(RunHV, 2)

rv0 = WaitDevOn(ChipString)

Call MsecDelay(0.2)

If CardResult <> 0 Then
    MsgBox "Read light On fail"
    End
End If

If rv0 = 1 Then
    OpenPipe
    rv0 = AU6435Set_Pad_Driving54(1)        'AU6435D51 set driving match customer setting
    ClosePipe
Else
    GoTo AU6435DLFResult
End If


CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
Call MsecDelay(0.02)

If (LightOn <> &H73) Then
    Tester.Print "LightON="; LightOn
    'Tester.Print "LightOFF="; LightOff
    Tester.Print "GPO_Fail or V18 out of range"
    UsbSpeedTestResult = GPO_FAIL
    rv0 = 3
    GoTo AU6435DLFResult
Else
    Tester.Print "V18-out Range PASS"
End If

Tester.Print "LBA="; LBA
 
ClosePipe

If rv0 = 1 Then
    rv0 = AU6435_CBWTest_New(0, 1, ChipString)
End If

If rv0 = 1 Then
    Call MsecDelay(0.02)
    rv0 = Read_SD30_Speed_AU6435(0, 0, 64, "4Bits")
    
    If rv0 <> 1 Then
        rv0 = 2
        Tester.Print "SD bus width Fail"
    End If
End If

If rv0 = 1 Then
    rv0 = Read_SD30_Mode_AU6435(0, 0, 64, "SDR")
    
    If rv0 <> 1 Then
        rv0 = 2
        Tester.Print "SD3.0 Mode Fail"
    End If
End If

'Tester.Print "rv0="; rv0
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      
TmpLBA = LBA
rv1 = 0
LBA = LBA + 199
            
If rv0 = 1 Then
    rv0 = CBWTest_New_128_Sector_PipeReady(0, rv0)  ' write
    ClosePipe
End If
                
ClosePipe

'If rv0 <> 1 Then
'    LBA = TmpLBA
'    GoTo AU6435DLFResult
'End If

'Next
LBA = TmpLBA
      
'=======================================================================================
                        
Call LabelMenu(0, rv0, 1)   ' no card test fail
     
Tester.Print rv0, " \\SDXC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

'===============================================
'  SDHC test
'================================================

Tester.Print "Force SD Card to SDHC Mode (Non-Ultra High Speed)"
OpenPipe
rv1 = ReInitial(0)
Call MsecDelay(0.02)
rv1 = AU6435ForceSDHC(rv0)
ClosePipe

If rv1 = 1 Then
    rv1 = AU6435_CBWTest_New(0, 1, ChipString)
End If


If rv1 = 1 Then
    rv1 = Read_SD30_Mode_AU6435(0, 0, 64, "Non-UHS")
    If rv1 <> 1 Then
        rv1 = 2
        Tester.Print "SD2.0 Mode Fail"
    End If
End If

ClosePipe

Call LabelMenu(1, rv1, rv0)

Tester.Print rv1, " \\SDHC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"


'===============================================
'  SMC Card test  : stop these test for card not enough
'================================================

'AU6433 has no SMC slot

rv2 = rv1   ' to complete the SMC asbolish

'===============================================
'  XD Card test
'================================================
CardResult = DO_WritePort(card, Channel_P1A, &H36) 'SD +XD
  
   
If CardResult <> 0 Then
    MsgBox "Set XD Card Detect On Fail"
    End
End If
  
  
Call MsecDelay(0.05)

CardResult = DO_WritePort(card, Channel_P1A, &H37)  'XD

Call MsecDelay(0.05)
 
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

CardResult = DO_WritePort(card, Channel_P1A, &H17)      'MS + XD

Call MsecDelay(0.05)

CardResult = DO_WritePort(card, Channel_P1A, &H1F)      'MS

OpenPipe
rv5 = ReInitial(0)
ClosePipe

rv5 = CBWTest_New(0, rv4, ChipString)

If rv5 = 1 Then
    rv5 = Read_MS_Speed_AU6435(0, 0, 64, "4Bits")
    
    If rv5 <> 1 Then
        rv5 = 2
        Tester.Print "MS bus width Fail"
    End If
End If
                
Call LabelMenu(31, rv5, rv4)
                
Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

ClosePipe


CardResult = DO_WritePort(card, Channel_P1A, &H3F)
Call MsecDelay(0.2)

If GetDeviceName_NoReply(ChipString) <> "" Then             'NBMD fail
    rv0 = 0
End If
                
               
AU6435DLFResult:

SetSiteStatus (HVDone)
Call WaitAnotherSiteDone(HVDone, 5)

SetSiteStatus (SiteUnknow)

CardResult = DO_WritePort(card, Channel_P1A, &H3F)   ' Close power
Call PowerSet2(0, "0.0", "0.7", 1, "0.0", "0.7", 1)


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

Public Sub AU6435GLE10TestSub()

'insert SD 3.0 Memory Card
'99/12/01 add V18-out detect
'100/2/18 Close Pad_Driving & tune pwr-off delay time
'100/3/1 Just revise LightON & LightOFF value (3.3V-Input)
'100/4/1 Skip HID function test
'100/6/3 AU6435D51 set driving, match customer setting
'100/8/4 Just modify Pin3 OS become open state, Pin2(Ext48in) skip
'2011/11/10: Purpose to solve V1.7 S/B Pin2(Ext48In) test coverage, this version is FT2 & using 48Mhz clock source
'2012/07/25: for Quanta RMA CSMC IC on smooth power curve will unknow issue
'This code copy from "AU6435ELF25TestSub" just modify SD card insert/remove 5 cycle

Dim TmpLBA As Long
Dim i As Integer
Dim DetectCount As Integer

If PCI7248InitFinish = 0 Then
    PCI7248ExistAU6254
    Call SetTimer_1ms
End If

OS_Result = 0
rv0 = 0
DetectCount = 0

'CardResult = DO_WritePort(card, Channel_P1C, &HFF)
'MsecDelay (0.02)

    
Tester.Print "AU6435GL Eng : Begin Test ..."
'==================================================================
'
'  this code come from AU6433DLF20TestSub
'  Add HID test function
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
Dim TmpCount As Integer
Dim InOutUnknowFlag As Boolean
OldChipName = ""
               

' initial condition

AU6371EL_SD = 1
AU6371EL_CF = 2
AU6371EL_XD = 8
AU6371EL_MS = 32
AU6371EL_MSP = 64
    
AU6371EL_BootTime = 0.6
     
    
If PCI7248InitFinish_Sync = 0 Then
  PCI7248Exist_P1C_Sync
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
InOutUnknowFlag = False

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

CardResult = DO_WritePort(card, Channel_P1A, &H3F)
Call MsecDelay(0.2)

If CardResult <> 0 Then
    MsgBox "Power off fail"
    End
End If

'Call MsecDelay(0.2)     'power on time

ChipString = "vid_058f"

 '================================================
CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
Call MsecDelay(0.02)
  
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
CardResult = DO_WritePort(card, Channel_P1A, &H3E)
      
If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If

Call PowerSet2(0, "5.0", "0.7", 1, "5.0", "0.7", 1)
Call MsecDelay(0.2)

SetSiteStatus (RunHV)
Call WaitAnotherSiteDone(RunHV, 2)

rv0 = WaitDevOn(ChipString)

If rv0 = 1 Then
    For TempCount = 1 To 5
        CardResult = DO_WritePort(card, Channel_P1A, &H3F)  'remove SD card
        WaitDevOFF (ChipString)
        Call MsecDelay(0.2)
        CardResult = DO_WritePort(card, Channel_P1A, &H3E)  'insert SD card
        If (WaitDevOn(ChipString) <> 1) Then
            rv0 = 0
            InOutUnknowFlag = True
            Exit For
        End If
    Next
End If

Call MsecDelay(0.2)

If CardResult <> 0 Then
    MsgBox "Read light On fail"
    End
End If

If rv0 = 1 Then
    OpenPipe
    rv0 = AU6435Set_Pad_Driving54(1)        'AU6435D51 set driving match customer setting
    ClosePipe
Else
    GoTo AU6435DLFResult
End If


CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
Call MsecDelay(0.02)

If (LightOn <> &H73) Then
    Tester.Print "LightON="; LightOn
    'Tester.Print "LightOFF="; LightOff
    Tester.Print "GPO_Fail or V18 out of range"
    UsbSpeedTestResult = GPO_FAIL
    rv0 = 3
    GoTo AU6435DLFResult
Else
    Tester.Print "V18-out Range PASS"
End If

Tester.Print "LBA="; LBA
 
ClosePipe

If rv0 = 1 Then
    rv0 = AU6435_CBWTest_New(0, 1, ChipString)
End If

If rv0 = 1 Then
    Call MsecDelay(0.02)
    rv0 = Read_SD30_Speed_AU6435(0, 0, 64, "4Bits")
    
    If rv0 <> 1 Then
        rv0 = 2
        Tester.Print "SD bus width Fail"
    End If
End If

If rv0 = 1 Then
    rv0 = Read_SD30_Mode_AU6435(0, 0, 64, "SDR")
    
    If rv0 <> 1 Then
        rv0 = 2
        Tester.Print "SD3.0 Mode Fail"
    End If
End If

'Tester.Print "rv0="; rv0
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      
TmpLBA = LBA
rv1 = 0
LBA = LBA + 199
            
If rv0 = 1 Then
    rv0 = CBWTest_New_128_Sector_PipeReady(0, rv0)  ' write
    ClosePipe
End If
                
ClosePipe

'If rv0 <> 1 Then
'    LBA = TmpLBA
'    GoTo AU6435DLFResult
'End If

'Next
LBA = TmpLBA
      
'=======================================================================================
                        
Call LabelMenu(0, rv0, 1)   ' no card test fail
     
Tester.Print rv0, " \\SDXC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

'===============================================
'  SDHC test
'================================================

Tester.Print "Force SD Card to SDHC Mode (Non-Ultra High Speed)"
OpenPipe
rv1 = ReInitial(0)
Call MsecDelay(0.02)
rv1 = AU6435ForceSDHC(rv0)
ClosePipe

If rv1 = 1 Then
    rv1 = AU6435_CBWTest_New(0, 1, ChipString)
End If


If rv1 = 1 Then
    rv1 = Read_SD30_Mode_AU6435(0, 0, 64, "Non-UHS")
    If rv1 <> 1 Then
        rv1 = 2
        Tester.Print "SD2.0 Mode Fail"
    End If
End If

ClosePipe

Call LabelMenu(1, rv1, rv0)

Tester.Print rv1, " \\SDHC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

'===============================================
'  SMC Card test  : stop these test for card not enough
'================================================

'AU6433 has no SMC slot

rv2 = rv1   ' to complete the SMC asbolish

'===============================================
'  XD Card test
'================================================
CardResult = DO_WritePort(card, Channel_P1A, &H36) 'SD +XD
  
   
If CardResult <> 0 Then
    MsgBox "Set XD Card Detect On Fail"
    End
End If
  
  
Call MsecDelay(0.05)

CardResult = DO_WritePort(card, Channel_P1A, &H37)  'XD

Call MsecDelay(0.05)
 
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

CardResult = DO_WritePort(card, Channel_P1A, &H17)      'MS + XD

Call MsecDelay(0.05)

CardResult = DO_WritePort(card, Channel_P1A, &H1F)      'MS

OpenPipe
rv5 = ReInitial(0)
ClosePipe

rv5 = CBWTest_New(0, rv4, ChipString)

If rv5 = 1 Then
    rv5 = Read_MS_Speed_AU6435(0, 0, 64, "4Bits")
    
    If rv5 <> 1 Then
        rv5 = 2
        Tester.Print "MS bus width Fail"
    End If
End If
                
Call LabelMenu(31, rv5, rv4)
                
Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

ClosePipe


CardResult = DO_WritePort(card, Channel_P1A, &H3F)
Call MsecDelay(0.2)

If GetDeviceName_NoReply(ChipString) <> "" Then             'NBMD fail
    rv0 = 0
End If
                
               
AU6435DLFResult:

SetSiteStatus (HVDone)
Call WaitAnotherSiteDone(HVDone, 5)

SetSiteStatus (SiteUnknow)

CardResult = DO_WritePort(card, Channel_P1A, &H3F)   ' Close power
Call PowerSet2(0, "0.0", "0.7", 1, "0.0", "0.7", 1)


If rv0 = UNKNOW Then
    UnknowDeviceFail = UnknowDeviceFail + 1
    If InOutUnknowFlag Then
        TestResult = "Bin5"
    Else
        TestResult = "Bin2"
    End If
ElseIf rv0 = WRITE_FAIL Then
    SDWriteFail = SDWriteFail + 1
    TestResult = "Bin2"
ElseIf rv0 = READ_FAIL Then
    SDReadFail = SDReadFail + 1
    TestResult = "Bin3"
ElseIf rv1 = WRITE_FAIL Then
    CFWriteFail = CFWriteFail + 1
    TestResult = "Bin3"
ElseIf rv1 = READ_FAIL Then
    CFReadFail = CFReadFail + 1
    TestResult = "Bin3"
ElseIf rv2 = WRITE_FAIL Or rv3 = WRITE_FAIL Then
    XDWriteFail = XDWriteFail + 1
    TestResult = "Bin4"
ElseIf rv2 = READ_FAIL Or rv3 = READ_FAIL Then
    XDReadFail = XDReadFail + 1
    TestResult = "Bin4"
 ElseIf rv4 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
    MSWriteFail = MSWriteFail + 1
    TestResult = "Bin4"
ElseIf rv4 = READ_FAIL Or rv5 = READ_FAIL Then
    MSReadFail = MSReadFail + 1
    TestResult = "Bin4"
ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
     TestResult = "PASS"
Else
    TestResult = "Bin2"
  
End If

End Sub

Public Sub AU6435ELF04TestSub()

'insert SD 3.0 Memory Card
'99/12/01 add V18-out detect
'100/2/18 Close Pad_Driving & tune pwr-off delay time
'100/3/1 Just revise LightON & LightOFF value (3.3V-Input)
'100/4/1 Skip HID function test
'100/6/3 AU6435D51 set driving, match customer setting
'100/8/4 Just modify Pin3 OS become open state, Pin2(Ext48in) skip
'2011/11/10: Purpose to solve V1.7 S/B Pin2(Ext48In) test coverage, this version is FT2 & using 48Mhz clock source
'2012/3/13: FT3 for V1.7 S/B, remove R40 & add 1x2 connector on "V33"


Dim TmpLBA As Long
Dim i As Integer
Dim DetectCount As Integer
Dim HV_Done_Flag As Boolean

If PCI7248InitFinish = 0 Then
    PCI7248ExistAU6254
    Call SetTimer_1ms
End If

OS_Result = 0
rv0 = 0
HV_Done_Flag = False

Routine_Label:

DetectCount = 0

If Not HV_Done_Flag Then
    Call PowerSet2(0, "3.6", "0.7", 1, "3.6", "0.7", 1)
    Tester.Print "AU6435EL : HV Begin Test ..."
Else
    Call PowerSet2(0, "3.2", "0.7", 1, "3.2", "0.7", 1)
    Tester.Print "AU6435EL : LV Begin Test ..."
End If

CardResult = DO_WritePort(card, Channel_P1C, &HFF)
MsecDelay (0.02)

'==================================================================
'
'  this code come from AU6433DLF20TestSub
'  Add HID test function
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
                     
                    
'If PCI7248InitFinish = 0 Then
'   PCI7248Exist
'End If
               
'  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
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
'Call MsecDelay(0.2)

If CardResult <> 0 Then
    MsgBox "Power off fail"
    End
End If

'Call MsecDelay(0.2)
 
CardResult = DO_WritePort(card, Channel_P1A, &H3F)
  
Call MsecDelay(0.2)     'power on time

ChipString = "vid"
       
'===============================================
'  SD Card test
'===============================================
  ' set SD card detect down
CardResult = DO_WritePort(card, Channel_P1A, &H3E)
      
If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If

Call MsecDelay(0.1)

rv0 = WaitDevOn(ChipString)

Call MsecDelay(0.2)

If CardResult <> 0 Then
    MsgBox "Read light On fail"
    End
End If

If rv0 = 1 Then
    OpenPipe
    rv0 = AU6435Set_Pad_Driving54(1)        'AU6435D51 set driving match customer setting
    ClosePipe
End If
                
CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
Call MsecDelay(0.02)

If (LightOn <> &H73) Then
    Tester.Print "LightON="; LightOn
    Tester.Print "LightOFF="; LightOff
    Tester.Print "GPO_Fail or V18 out of range"
    UsbSpeedTestResult = GPO_FAIL
    rv0 = 3
    GoTo AU6435DLFResult
Else
    Tester.Print "V18-out Range PASS"
End If

Tester.Print "LBA="; LBA
 
ClosePipe

If rv0 = 1 Then
    rv0 = AU6435_CBWTest_New(0, 1, ChipString)
End If

If rv0 = 1 Then
    Call MsecDelay(0.02)
    rv0 = Read_SD30_Speed_AU6435(0, 0, 64, "4Bits")
    
    If rv0 <> 1 Then
        rv0 = 2
        Tester.Print "SD bus width Fail"
    End If
End If

If rv0 = 1 Then
    rv0 = Read_SD30_Mode_AU6435(0, 0, 64, "SDR")
    
    If rv0 <> 1 Then
        rv0 = 2
        Tester.Print "SD3.0 Mode Fail"
    End If
End If

'Tester.Print "rv0="; rv0
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      
TmpLBA = LBA
rv1 = 0
LBA = LBA + 199
            
If rv0 = 1 Then
    rv0 = CBWTest_New_128_Sector_PipeReady(0, rv0)  ' write
    ClosePipe
End If
                
ClosePipe

LBA = TmpLBA
                                             
Call LabelMenu(0, rv0, 1)   ' no card test fail
     
Tester.Print rv0, " \\SDXC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

                
'===============================================
'  SDHC test
'================================================

Tester.Print "Force SD Card to SDHC Mode (Non-Ultra High Speed)"
OpenPipe
rv1 = ReInitial(0)
Call MsecDelay(0.02)
rv1 = AU6435ForceSDHC(rv0)
ClosePipe

If rv1 = 1 Then
    rv1 = AU6435_CBWTest_New(0, 1, ChipString)
End If


If rv1 = 1 Then
    rv1 = Read_SD30_Mode_AU6435(0, 0, 64, "Non-UHS")
    If rv1 <> 1 Then
        rv1 = 2
        Tester.Print "SD2.0 Mode Fail"
    End If
End If

ClosePipe

Call LabelMenu(1, rv1, rv0)

Tester.Print rv1, " \\SDHC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                
'===============================================
'  SMC Card test  : stop these test for card not enough
'================================================

'AU6433 has no SMC slot

rv2 = rv1   ' to complete the SMC asbolish
               
'===============================================
'  XD Card test
'================================================
CardResult = DO_WritePort(card, Channel_P1A, &H36) 'SD +XD
  
   
If CardResult <> 0 Then
    MsgBox "Set XD Card Detect On Fail"
    End
End If
                  
Call MsecDelay(0.05)

CardResult = DO_WritePort(card, Channel_P1A, &H37)  'XD

Call MsecDelay(0.05)
 
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

CardResult = DO_WritePort(card, Channel_P1A, &H17)      'MS + XD

Call MsecDelay(0.05)

CardResult = DO_WritePort(card, Channel_P1A, &H1F)      'MS

OpenPipe
rv5 = ReInitial(0)
ClosePipe

rv5 = CBWTest_New(0, rv4, ChipString)
                
If rv5 = 1 Then
    rv5 = Read_MS_Speed_AU6435(0, 0, 64, "4Bits")
    
    If rv5 <> 1 Then
        rv5 = 2
        Tester.Print "MS bus width Fail"
    End If
End If
          
ClosePipe
Call LabelMenu(31, rv5, rv4)

Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"



CardResult = DO_WritePort(card, Channel_P1A, &H3F)
Call MsecDelay(0.2)
                
If GetDeviceName_NoReply(ChipString) <> "" Then             'NBMD fail
    rv0 = 0
End If
                

AU6435DLFResult:
                
    CardResult = DO_WritePort(card, Channel_P1A, &H3F)   ' Close power
    Call PowerSet2(0, "0.0", "0.7", 1, "0.0", "0.7", 1)
    
    If HV_Done_Flag = False Then
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
        HV_Done_Flag = True
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

Public Sub AU6435ELS11TestSub()

'insert SD 3.0 Memory Card
'99/12/01 add V18-out detect
'100/2/18 Close Pad_Driving & tune pwr-off delay time
'100/3/1 Just revise LightON & LightOFF value (3.3V-Input)

Dim TmpLBA As Long
Dim i As Integer
Dim Detect_Counter As Integer

Call PowerSet2(1, "5.0", "0.7", 1, "5.0", "0.7", 1)

Tester.Print "AU6435EL :5V Begin Test ..."
'==================================================================
'
'  this code come from AU6433DLF20TestSub
'  Add HID test function
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
               
               '  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
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
                
                If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                End If
                
                'Call MsecDelay(0.2)
                 
                CardResult = DO_WritePort(card, Channel_P1A, &H3F)
                  
                Call MsecDelay(0.3)     'power on time
                
                ChipString = "vid"
                 
                
                
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
                CardResult = DO_WritePort(card, Channel_P1A, &H3E)
                      
                If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect Down Fail"
                    End
                End If
                
                Call MsecDelay(0.3)
                
                rv0 = WaitDevOn(ChipString)
                
                Call MsecDelay(0.2)
                
                If CardResult <> 0 Then
                    MsgBox "Read light On fail"
                    End
                End If
                      
                ClosePipe
                
                If rv0 = 1 Then
                    rv0 = AU6435_CBWTest_New(0, 1, ChipString)
                End If
                
                If rv0 = 1 Then
                    Call MsecDelay(0.02)
                    rv0 = Read_SD30_Speed_AU6435(0, 0, 64, "4Bits")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
                    End If
                End If
                
                If rv0 = 1 Then
                    rv0 = Read_SD30_Mode_AU6435(0, 0, 64, "SDR")
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD3.0 Mode Fail"
                    End If
                End If
                
                ClosePipe
                      
                Tester.Print "rv0="; rv0
                     
'=======================================================================================
    'SD R / W
'=======================================================================================
                      
                TmpLBA = LBA
                rv1 = 0
                LBA = LBA + 199
                            
                ClosePipe
                
                If rv0 = 1 Then
                    rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                    ClosePipe
                End If
                                
                If rv1 <> 1 Then
                    LBA = TmpLBA
                    GoTo AU6435DLFResult
                End If
                
                'Next
                LBA = TmpLBA
                      
'=======================================================================================
                        
                    
                     
                Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                Tester.Print rv0, " \\SDXC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                '===============================================
                '  SDHC test
                '================================================
                
                Tester.Print "Force SD Card to SDHC Mode (Non-Ultra High Speed)"
                OpenPipe
                rv1 = ReInitial(0)
                Call MsecDelay(0.02)
                rv1 = AU6435ForceSDHC(rv0)
                ClosePipe
                
                If rv1 = 1 Then
                    rv1 = AU6435_CBWTest_New(0, 1, ChipString)
                End If
                
                
                If rv1 = 1 Then
                    rv1 = Read_SD30_Mode_AU6435(0, 0, 64, "Non-UHS")
                    If rv1 <> 1 Then
                        rv1 = 2
                        Tester.Print "SD2.0 Mode Fail"
                    End If
                End If
                
                ClosePipe
                
                Call LabelMenu(1, rv1, rv0)
            
                Tester.Print rv1, " \\SDHC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
              
                'AU6433 has no SMC slot
               
                rv2 = rv1   ' to complete the SMC asbolish
               
                'rv3 = 1
                'rv4 = 1
                'rv5 = 1
                'GoTo AU6435tmpResult
                '===============================================
                '  XD Card test
                '================================================
                CardResult = DO_WritePort(card, Channel_P1A, &H36) 'SD +XD
                  
                   
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect On Fail"
                    End
                End If
                  
                  
                Call MsecDelay(0.05)
                
                CardResult = DO_WritePort(card, Channel_P1A, &H37)  'XD
                
                Call MsecDelay(0.05)
                 
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
              
                
                
                CardResult = DO_WritePort(card, Channel_P1A, &H37)      'MS + XD
              
                Call MsecDelay(0.05)
               
                CardResult = DO_WritePort(card, Channel_P1A, &H1F)      'MS
               
                OpenPipe
                rv5 = ReInitial(0)
                ClosePipe
                
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                    If rv5 = 1 Then
                        rv5 = Read_MS_Speed_AU6435(0, 0, 64, "4Bits")
                        
                        If rv5 <> 1 Then
                            rv5 = 2
                            Tester.Print "MS bus width Fail"
                        End If
                    End If
                
                Call LabelMenu(31, rv5, rv4)
                
                Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                ClosePipe
                
       
                CardResult = DO_WritePort(card, Channel_P1A, &H3F)
                Call MsecDelay(0.2)
                
                If GetDeviceName_NoReply(ChipString) <> "" Then             'NBMD fail
                    rv0 = 0
                    GoTo AU6435DLFResult
                End If
                
                '=================================================================================
                ' HID mode and reader mode ---> compositive device
                If rv5 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &HAF) '  pwr off  for HID mode
                    'result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                    
                    Call MsecDelay(0.2)
          
                    CardResult = DO_WritePort(card, Channel_P1A, &H6E) ' HID mode   'PID_6466
                            
                    Call MsecDelay(0.3)
                    rv5 = WaitDevOn(ChipString)
                    Call MsecDelay(0.1)
                    Detect_Counter = 0
                    
                    LightOn = 0
                    
                    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
          
                End If
                    
                If (rv5 = 1) And ((LightOn <> &H13) Or (LightOff <> &H93)) Then
                    Tester.Print "LightON="; LightOn
                    Tester.Print "LightOFF="; LightOff
                    Tester.Print "GPO_Fail or V18 out of range"
                    UsbSpeedTestResult = GPO_FAIL
                    rv0 = 3
                    GoTo AU6435DLFResult
                Else
                    Tester.Print "V18-out Range PASS"
                End If
                
                If rv4 = 1 Then
                ' code begin
                     
                    Tester.Cls
                    Tester.Print "keypress test begin---------------"
                    Dim ReturnValue As Byte
HIDRetest:
                    DeviceHandle = &HFFFF  'invalid handle initial value
                     
                    ReturnValue = fnGetDeviceHandle(DeviceHandle)
                    Tester.Print ReturnValue; Space(5); ' 1: pass the other refer btnstatus.h
                    Tester.Print "DeviceHandle="; DevicehHandle
                     
                    If ReturnValue <> 1 Then
                        rv0 = UNKNOW       '---> HID mode unknow device mode
                        Call LabelMenu(0, rv0, 1)
                        Tester.Label9.Caption = "HID mode unknow device"
                        fnFreeDeviceHandle (DeviceHandle)
                        GoTo AU6435DLFResult
                    End If
                     
                    '=======================
                    '  key press test, it will return 10 when key up, GPI 6 must do low go hi action
                    '========================

                
                    Do
                        CardResult = DO_WritePort(card, Channel_P1A, &H6E) 'GPI6 : bit 6: pull high
                        Sleep (200)
                        CardResult = DO_WritePort(card, Channel_P1A, &H2E)  ' GPI6 : bit 6: pull low
                        Sleep (500)
                       
                        ReturnValue = fnInquiryBtnStatus(DeviceHandle)
                        Tester.Print i; Space(5); "Key press value="; ReturnValue
                        i = i + 1
                    Loop While i < 3 And ReturnValue <> 10
                     
                    If (ReturnValue = 12) And ((i = 3) Or (i = 4)) Then
                         
                        GoTo HIDRetest
                    End If
                    
                    If ReturnValue <> 10 Then
                     
                        rv1 = 2
                        Call LabelMenu(1, rv1, rv0)
                        Tester.Label9.Caption = "KeyPress Fail"
                       
                    End If
                              
                End If
AU6435tmpResult:
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print "LBA="; LBA
                 
               
AU6435DLFResult:
                
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
