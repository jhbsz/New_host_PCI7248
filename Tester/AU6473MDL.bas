Attribute VB_Name = "AU6473MDL"
Public Sub AU6473TestSub()

    If ChipName = "AU6473CLF20" Then
        Call AU6473CLF20TestSub
    End If
    
    If ChipName = "AU6473CLF21" Then
        Call AU6473CLF21TestSub
    End If
    
    If ChipName = "AU6473BLF20" Then
        Call AU6473BLF20TestSub
    End If
    
    If ChipName = "AU6473CLF31" Then
        Call AU6473CLF31TestSub
    End If

    If ChipName = "AU6429FLF20" Then
        Call AU6429FLF20TestSub
    End If
    
    If ChipName = "AU6425DLF20" Then
        Call AU6425DLF20TestSub
    End If
    
    If ChipName = "AU6427ELF20" Then
        Call AU6427ELF20TestSub
    End If
    
    If ChipName = "AU6427GLF20" Then
        Call AU6427GLF20TestSub
    End If
End Sub

Public Sub AU6473CLF20TestSub()

Dim TmpLBA As Long
Dim i As Integer
      
    Tester.Print "AU6473CL Test Begin ..."
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
            
            
            'If GetDeviceName(ChipString) <> "" Then
            '    rv0 = 0
            '    GoTo AU6371ELResult
            'End If
                 
        '================================================
        '    CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
        '
        '    If CardResult <> 0 Then
        '        MsgBox "Read light off fail"
        '        End
        '    End If
                   
                 
               
        '************************************************
        '*               SD Card test                   *
        '************************************************
        
            'Call MsecDelay(0.3)
             
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
                      
                rv0 = CBWTest_New(1, 1, ChipString)
                            
                'Call MsecDelay(0.1)
                
                If rv0 = 1 Then
                    rv0 = Read_SD_Speed_AU6473(0, 1, 60, "4Bits")
                           
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
                    End If
                End If
                      
                ClosePipe
                
                Tester.Print "rv0="; rv0
                     
                'If rv0 <> 0 Then
                '    If LightON <> &HBF Or LightOFF <> &HFF Then
                '        Tester.Print "LightON="; LightON
                '        Tester.Print "LightOFF="; LightOFF
                '        UsbSpeedTestResult = GPO_FAIL
                '        rv0 = 3
                '    End If
                'End If
                    
                     
        '=======================================================================================
        'SD R / W
        '=======================================================================================
                Call MsecDelay(0.1)
                
                TmpLBA = LBA
                rv1 = 0
                LBA = LBA + 199
                            
                ClosePipe
                
                rv1 = CBWTest_New_128_Sector_AU6377(1, 1)  ' write
                    If rv1 <> 1 Then
                        LBA = TmpLBA
                        GoTo AU6371ELResult
                    End If
                
                LBA = TmpLBA
                      
                Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                
        
        '************************************************
        '*               CF Card test                   *
        '************************************************
                   
                rv1 = rv0  'AU6473 no CF slot
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                'Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
        
        '************************************************
        '*               SMC Card test                   *
        '************************************************
              
                rv2 = rv1   'AU6473 no SMC slot
                 
                
              
        '************************************************
        '*               XD Card test                   *
        '************************************************
                
                rv3 = rv2   'AU6473 no XD slot
                
        '************************************************
        '*               MS Card test                   *
        '************************************************
                   
                     
                rv4 = rv3  'AU6473 has no MS slot pin
               
                'Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
        
        '************************************************
        '*               MS PRO Card test               *
        '************************************************
              
                
                
                CardResult = DO_WritePort(card, Channel_P1A, &H7F)
              
                Call MsecDelay(0.3)
               
                CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                 
                Call MsecDelay(0.1)
                'OpenPipe
                '    rv5 = ReInitial(0)
                'ClosePipe
                
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                
                If rv5 = 1 Then
                    rv5 = Read_MS_Speed_AU6473(0, 0, 60, "4Bits")
                    
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

Public Sub AU6473CLF21TestSub()

Dim TmpLBA As Long
Dim i As Integer
      
    Tester.Print "AU6473CL Test Begin ..."
'==================================================================
'
'  this code come from AU6473CLF21TestSub
'  purpose to solve when test fail always power-on issue.
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
            
            
            'If GetDeviceName(ChipString) <> "" Then
            '    rv0 = 0
            '    GoTo AU6371ELResult
            'End If
                 
        '================================================
        '    CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
        '
        '    If CardResult <> 0 Then
        '        MsgBox "Read light off fail"
        '        End
        '    End If
                   
                 
               
        '************************************************
        '*               SD Card test                   *
        '************************************************
        
            'Call MsecDelay(0.3)
             
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
                
                rv0 = CBWTest_New(1, 1, ChipString)
                            
                'Call MsecDelay(0.1)
                
                If rv0 = 1 Then
                    rv0 = Read_SD_Speed_AU6473(0, 1, 60, "4Bits")
                           
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
                    End If
                End If
                      
                ClosePipe
                
                Tester.Print "rv0="; rv0
                     
                'If rv0 <> 0 Then
                '    If LightON <> &HBF Or LightOFF <> &HFF Then
                '        Tester.Print "LightON="; LightON
                '        Tester.Print "LightOFF="; LightOFF
                '        UsbSpeedTestResult = GPO_FAIL
                '        rv0 = 3
                '    End If
                'End If
                    
                     
        '=======================================================================================
        'SD R / W
        '=======================================================================================
                Call MsecDelay(0.1)
                
                TmpLBA = LBA
                rv1 = 0
                LBA = LBA + 199
                            
                ClosePipe
                
                If rv0 = 1 Then
                    rv1 = CBWTest_New_128_Sector_AU6377(1, 1)  ' write
                
                
                    If rv1 <> 1 Then
                        LBA = TmpLBA
                        GoTo AU6371ELResult
                    End If
                End If
                
                LBA = TmpLBA
                      
                Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                
        
        '************************************************
        '*               CF Card test                   *
        '************************************************
                   
                rv1 = rv0  'AU6473 no CF slot
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                'Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
        
        '************************************************
        '*               SMC Card test                   *
        '************************************************
              
                rv2 = rv1   'AU6473 no SMC slot
                 
                
              
        '************************************************
        '*               XD Card test                   *
        '************************************************
                
                rv3 = rv2   'AU6473 no XD slot
                
        '************************************************
        '*               MS Card test                   *
        '************************************************
                   
                     
                rv4 = rv3  'AU6473 has no MS slot pin
               
                'Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
        
        '************************************************
        '*               MS PRO Card test               *
        '************************************************
              
                
                
                CardResult = DO_WritePort(card, Channel_P1A, &H7F)
              
                Call MsecDelay(0.3)
               
                CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                 
                Call MsecDelay(0.1)
                'OpenPipe
                '    rv5 = ReInitial(0)
                'ClosePipe
                
                
                    rv5 = CBWTest_New(0, rv4, ChipString)
               
                If rv5 = 1 Then
                    rv5 = Read_MS_Speed_AU6473(0, 0, 60, "4Bits")
                    
                    If rv5 <> 1 Then
                        rv5 = 2
                        Tester.Print "MS bus width Fail"
                    End If
                End If
                
                Call LabelMenu(31, rv5, rv4)
                
                Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                 
               

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
Public Sub AU6429FLF20TestSub()

Dim TmpLBA As Long
Dim i As Integer
      
    Tester.Print "AU6429FL Test Begin ..."
'==================================================================
'
'  this code come from AU6473CLF21TestSub
'  purpose to solve when test fail always power-on issue.
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
            ChipString = "vid"
            
            CardResult = DO_WritePort(card, Channel_P1A, &H3F)
            Call MsecDelay(0.2)
            
            rv0 = 1
            
            If CardResult <> 0 Then
                MsgBox "Power off fail"
                End
            End If
                 
            If GetDeviceName(ChipString) <> "" Then
                rv0 = 0
                GoTo AU6429FLResult
            End If
            
        '================================================
            
            If CardResult <> 0 Then
                MsgBox "Read light off fail"
                End
            End If
                   
            CardResult = DO_WritePort(card, Channel_P1A, &H3E)  'ENA, MS1 pwr-on
                  
            Call MsecDelay(1.3)        'power on time
                 
            CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
            
                'AU6429GFL has no sd slot
            
        '************************************************
        '*               SD Card test                   *
        '************************************************
        
            'Call MsecDelay(0.3)
             
            '===========================================
            'NO card test
            '============================================
  
                If rv0 <> 0 Then
                    If LightOn <> &HFE Then
                        Tester.Print "LightON="; LightOn
                        UsbSpeedTestResult = GPO_FAIL
                        rv0 = 3
                    End If
                End If
                         
        '=======================================================================================
        'SD R / W
        '=======================================================================================
                
                TmpLBA = LBA
                rv1 = 0
                'LBA = LBA + 199
                
                'LBA = TmpLBA
                      
                Call LabelMenu(0, rv0, 1)   ' no card test fail
                 
        '************************************************
        '*               CF Card test                   *
        '************************************************
                   
                rv1 = rv0  'AU6429GFL has no CF slot
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                'Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
        
        '************************************************
        '*               SMC Card test                   *
        '************************************************
              
                rv2 = rv1   'AU6429GFL has no smc slot
                 
                
              
        '************************************************
        '*               XD Card test                   *
        '************************************************
                
                rv3 = rv2   'AU6429GFL has no xd slot
                
        '************************************************
        '*               MS Card test                   *
        '************************************************
                   
                     
                rv4 = rv3   'AU6429GFL has no MS slot
                
                'Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
        
        '************************************************
        '*               MS PRO Card test               *
        '************************************************
              
                
                
                'rv5 = CBWTest_New(0, rv4, ChipString)
               
                rv5 = CBWTest_New(0, rv4, ChipString)   'LUN0_MSpro
                Call MsecDelay(0.05)
                    
                 If rv5 = 1 Then
                    rv5 = Read_MS_Speed_AU6473(LBA, 0, 60, "4Bits")
                    Call MsecDelay(0.05)
                        
                    If rv5 <> 1 Then
                        rv5 = 2
                        Tester.Print "MS bus width Fail"
                    Else
                        Tester.Print "LUN0 MSpro PASS"
                    End If
                    
                End If
                
                Call LabelMenu(3, rv5, rv4)
                
                ClosePipe
                
                Call MsecDelay(0.1)
                LBA = LBA + 1
                
                If rv5 = 1 Then
                    rv5 = CBWTest_New(1, rv4, ChipString)   'LUN1_MSpro
                    Call MsecDelay(0.05)
                    
                    If rv5 = 1 Then
                        rv5 = Read_MS_Speed_AU6473(LBA, 1, 60, "4Bits")
                        Call MsecDelay(0.05)
                        
                        If rv5 <> 1 Then
                            rv5 = 2
                            Tester.Print "MS bus width Fail"
                        Else
                            Tester.Print "LUN1 MSpro PASS"
                        End If
                    End If
                End If
            
                
                ClosePipe
                Call LabelMenu(31, rv5, rv4)
                
                Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
               

AU6429FLResult:
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
Public Sub AU6425DLF20TestSub()

Dim TmpLBA As Long
Dim i As Integer
      
    Tester.Print "AU6429FL Test Begin ..."
'==================================================================
'
'  this code come from AU6473CLF21TestSub
'  purpose to solve when test fail always power-on issue.
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
            
            If GetDeviceName(ChipString) <> "" Then
                rv0 = 0
                GoTo AU6425FLResult
            End If
                
            CardResult = DO_WritePort(card, Channel_P1A, &H7F)
            Call MsecDelay(0.9)
                
        '=========================================
        '    POWER on
        '=========================================
            ChipString = "vid"
            
            CardResult = DO_WritePort(card, Channel_P1A, &H7C)  'SD0, SD1 pwr-on
            Call MsecDelay(0.3)
            
            If CardResult <> 0 Then
                MsgBox "Power off fail"
                End
            End If
                 
        '================================================
            
            If CardResult <> 0 Then
                MsgBox "Read light off fail"
                End
            End If
                   
            CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
            
                'AU6429GFL has no sd slot
            
        '************************************************
        '*               SD Card test                   *
        '************************************************
        
            'Call MsecDelay(0.3)
             
            '===========================================
            'NO card test
            '============================================
  
                If rv0 <> 0 Then
                    If LightOn <> &HFE Then
                        Tester.Print "LightON="; LightOn
                        UsbSpeedTestResult = GPO_FAIL
                        rv0 = 3
                    End If
                End If
                         
        '=======================================================================================
        'SD R / W
        '=======================================================================================
                
                TmpLBA = LBA
                                     
                rv0 = CBWTest_New(0, 1, ChipString)
                
                If rv0 = 1 Then
                    rv0 = Read_SD_Speed_AU6473(TmpLBA, 0, 60, "8Bits")   'Lun0 SD
                    Tester.Print "LUN0 MMC= "; rv0
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
                    ElseIf rv0 = 1 Then
                        
                        rv0 = CBWTest_New_128_Sector_AU6377(0, rv0)  ' write          LUN0_SD
                        
                        If rv0 <> 1 Then
                            rv0 = 2
                            LBA = TmpLBA
                        End If
                    End If
                    
                    ClosePipe
                    Call MsecDelay(0.2)
                    
                    If rv0 = 1 Then
                        rv0 = CBWTest_New(1, rv0, ChipString)
                        
                        If rv0 = 1 Then
                            rv0 = Read_SD_Speed_AU6473(TmpLBA, 1, 60, "4Bits")   'Lun1 SD
                            Tester.Print "LUN1 SD= "; rv0
                        End If
                        
                        If rv0 <> 1 Then
                            rv0 = 2
                            Tester.Print "SD bus width Fail"
                        Else
                            rv0 = CBWTest_New_128_Sector_AU6377(1, rv0)  ' write          LUN1_SD
                            
                            If rv0 <> 1 Then
                                rv0 = 2
                                LBA = TmpLBA
                            End If
                        
                        End If
                    End If
                End If
                      
                ClosePipe
                
                Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                                      
                                      
                                      
                Call LabelMenu(0, rv0, 1)   ' no card test fail
                 
        '************************************************
        '*               CF Card test                   *
        '************************************************
                   
                rv1 = rv0  'AU6425DL has no CF slot
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                'Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
        
        '************************************************
        '*               SMC Card test                   *
        '************************************************
              
                rv2 = rv1   'AU6425DL has no smc slot
                 
                
              
        '************************************************
        '*               XD Card test                   *
        '************************************************
                
                rv3 = rv2   'AU6425DL has no xd slot
                
        '************************************************
        '*               MS Card test                   *
        '************************************************
                   
                     
                rv4 = rv3   'AU6425DL has no MS slot
                
                'Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
        
        '************************************************
        '*               MS PRO Card test               *
        '************************************************
              
                
                
                rv5 = rv4   'AU6425DL has no MSpro slot
               
               

AU6425FLResult:
                CardResult = DO_WritePort(card, Channel_P1A, &HF0)   ' Close power
                 
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
Public Sub AU6427ELF20TestSub()

Dim TmpLBA As Long
Dim i As Integer
      
    Tester.Print "AU6427EL Test Begin ..."
'==================================================================
'
'  this code come from AU6473CLF21TestSub
'  purpose to solve when test fail always power-on issue.
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
            
            If GetDeviceName(ChipString) <> "" Then
                rv0 = 0
                GoTo AU6425FLResult
            End If
                
            CardResult = DO_WritePort(card, Channel_P1A, &H7F)
            Call MsecDelay(0.6)
                
        '=========================================
        '    POWER on
        '=========================================
            ChipString = "vid"
            
            CardResult = DO_WritePort(card, Channel_P1A, &H7C)  'SD0, MS1 pwr-on
            Call MsecDelay(0.2)
            
            If CardResult <> 0 Then
                MsgBox "Power off fail"
                End
            End If
                 
        '================================================
            
            If CardResult <> 0 Then
                MsgBox "Read light off fail"
                End
            End If
                   
            CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
            
            
        '************************************************
        '*               SD Card test                   *
        '************************************************
        
            'Call MsecDelay(0.3)
             
            '===========================================
            'NO card test
            '============================================
  
                If rv0 <> 0 Then
                    If LightOn <> &HFE Then
                        Tester.Print "LightON="; LightOn
                        UsbSpeedTestResult = GPO_FAIL
                        rv0 = 3
                    End If
                End If
                         
        '=======================================================================================
        'SD R / W
        '=======================================================================================
                
                TmpLBA = LBA
                                     
                rv0 = CBWTest_New(0, 1, ChipString)
                
                If rv0 = 1 Then
                    rv0 = Read_SD_Speed_AU6473(TmpLBA, 0, 60, "8Bits")   'Lun0 SD
                    Tester.Print "LUN0 MMC= "; rv0
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
                    ElseIf rv0 = 1 Then
                        
                        rv0 = CBWTest_New_128_Sector_AU6377(0, rv0)  ' write          LUN0_SD
                        
                        If rv0 <> 1 Then
                            rv0 = 2
                            LBA = TmpLBA
                        End If
                    End If
                    
                    ClosePipe
                    Call MsecDelay(0.01)
                    
                End If
                      
                ClosePipe
                
                Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                                      
                                      
                                      
                Call LabelMenu(0, rv0, 1)   ' no card test fail
                 
        '************************************************
        '*               CF Card test                   *
        '************************************************
                   
                rv1 = rv0  'AU6425DL has no CF slot
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                'Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
        
        '************************************************
        '*               SMC Card test                   *
        '************************************************
              
                rv2 = rv1   'AU6425DL has no smc slot
                 
                
              
        '************************************************
        '*               XD Card test                   *
        '************************************************
                
                rv3 = rv2   'AU6425DL has no xd slot
                
        '************************************************
        '*               MS Card test                   *
        '************************************************
                   
                     
                rv4 = rv3   'AU6425DL has no MS slot
                
                'Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
        
        '************************************************
        '*               MS PRO Card test               *
        '************************************************
            If rv4 = 1 Then
                
                rv5 = CBWTest_New(1, rv4, ChipString)   'LUN1_MSpro
                    
                If rv5 = 1 Then
                    rv5 = Read_MS_Speed_AU6473(TmpLBA, 1, 60, "4Bits")
                         
                    If rv5 <> 1 Then
                        rv5 = 2
                        Tester.Print "MS bus width Fail"
                        Call LabelMenu(1, rv5, rv4)
                    Else
                        Tester.Print "LUN1 MSpro PASS"
                    End If
                End If
            End If
                
            ClosePipe
            Call LabelMenu(31, rv5, rv4)
                
            Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
               

AU6425FLResult:
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
Public Sub AU6427GLF20TestSub()

Dim TmpLBA As Long
Dim i As Integer
      
    Tester.Print "AU6427GL Test Begin ..."
'==================================================================
'
'  this code come from AU6473ELF20TestSub
'  purpose to solve when test fail always power-on issue.
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
            
            'CardResult = DO_WritePort(card, Channel_P1A, &H80)
            'Call MsecDelay(0.2)
            
            If GetDeviceName(ChipString) <> "" Then
                rv0 = 0
                Tester.Print "PowerOFF Fail"
            Else
                rv0 = 1
            End If
            
            CardResult = DO_WritePort(card, Channel_P1A, &H7F)
            Call MsecDelay(0.3)
            CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
            Call MsecDelay(0.02)
        '=========================================
        '    POWER on
        '=========================================
            ChipString = "vid"
            
            CardResult = DO_WritePort(card, Channel_P1A, &H7C)  'SD1, MS0 pwr-on
            
            If CardResult <> 0 Then
                MsgBox "Power off fail"
                End
            End If
            
            If rv0 = 1 Then
                Call MsecDelay(0.2)
                rv0 = WaitDevOn(ChipString)
                Call MsecDelay(0.1)
            End If
            
        '================================================
                   
            CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
            
            If CardResult <> 0 Then
                MsgBox "Read light off fail"
                End
            End If
            
            If rv0 = 1 Then
                If (LightOn <> &HFE) Or (LightOff <> &HFF) Then
                        Tester.Print "LightON="; LightOn
                        Tester.Print "LightOff="; LightOff
                        UsbSpeedTestResult = GPO_FAIL
                        rv0 = 3
                End If
            End If
            
        '************************************************
        '*               SD Card test                   *
        '************************************************
                         
        '=======================================================================================
        'SD R / W
        '=======================================================================================
                
                TmpLBA = LBA
                                     
                If rv0 = 1 Then
                    rv0 = CBWTest_New(1, 1, ChipString)
                End If
                
                If rv0 = 1 Then
                    rv0 = Read_SD_Speed_AU6473(TmpLBA, 1, 60, "4Bits")   'Lun1 SD
                    Tester.Print "LUN1 SD/MMC= "; rv0
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
                    ElseIf rv0 = 1 Then
                        
                        rv0 = CBWTest_New_128_Sector_AU6377(1, rv0)  ' write          LUN1_SD
                        
                        If rv0 <> 1 Then
                            rv0 = 2
                            LBA = TmpLBA
                        End If
                    End If
                    
                    ClosePipe
                    
                End If
                      
                ClosePipe
                
                Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                                      
                                      
                                      
                Call LabelMenu(0, rv0, 1)   ' no card test fail
                 
        '************************************************
        '*               CF Card test                   *
        '************************************************
                   
                rv1 = rv0  'AU6425DL has no CF slot
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                'Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
        
        '************************************************
        '*               SMC Card test                   *
        '************************************************
              
                rv2 = rv1   'AU6425DL has no smc slot
                 
                
              
        '************************************************
        '*               XD Card test                   *
        '************************************************
                
                rv3 = rv2   'AU6425DL has no xd slot
                
        '************************************************
        '*               MS Card test                   *
        '************************************************
                   
                     
                rv4 = rv3   'AU6425DL has no MS slot
                
                'Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
        
        '************************************************
        '*               MS PRO Card test               *
        '************************************************
            If rv4 = 1 Then
                
                rv5 = CBWTest_New(0, rv4, ChipString)   'LUN0_MSpro
                    
                If rv5 = 1 Then
                    rv5 = Read_MS_Speed_AU6473(TmpLBA, 0, 60, "4Bits")
                         
                    If rv5 <> 1 Then
                        rv5 = 2
                        Tester.Print "MS bus width Fail"
                        Call LabelMenu(1, rv5, rv4)
                    End If
                End If
            End If
                
            ClosePipe
            Call LabelMenu(31, rv5, rv4)
                
            Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            
AU6425FLResult:
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
Public Sub AU6473CLF31TestSub()

Dim TmpLBA As Long
Dim i As Integer

If PCI7248InitFinish = 0 Then
    PCI7248ExistAU6254
    Call SetTimer_1ms
End If

OS_Result = 0
rv0 = 0

CardResult = DO_WritePort(card, Channel_P1C, &H0)   'Set Switch connect to OS Board
                 
Call MsecDelay(0.05)

OpenShortTest_Result


If OS_Result <> 1 Then
    rv0 = 0                 'OS Fail
    GoTo AU6371ELResult
End If

CardResult = DO_WritePort(card, Channel_P1C, &HFF)   'Set Switch connect to FT Module Board
Call MsecDelay(0.05)

    Tester.Print "AU6473CL Test Begin ..."
'==================================================================
'
'  this code come from AU6473CLF21TestSub
'  purpose to solve when test fail always power-on issue.
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
            Call MsecDelay(0.1)
            
            If CardResult <> 0 Then
                MsgBox "Power off fail"
                End
            End If
                 
                 
            CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
            Call MsecDelay(1#)      'power on time
            
            ChipString = "vid"
            
            
            'If GetDeviceName(ChipString) <> "" Then
            '    rv0 = 0
            '    GoTo AU6371ELResult
            'End If
                 
        '================================================
        '    CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
        '
        '    If CardResult <> 0 Then
        '        MsgBox "Read light off fail"
        '        End
        '    End If
                   
                 
               
        '************************************************
        '*               SD Card test                   *
        '************************************************
        
            'Call MsecDelay(0.3)
             
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
                Call MsecDelay(0.01)
                
                If CardResult <> 0 Then
                    MsgBox "Read light On fail"
                    End
                End If
                           
                
                rv0 = CBWTest_New(1, 1, ChipString)
                            
                'Call MsecDelay(0.1)
                
                If rv0 = 1 Then
                    rv0 = Read_SD_Speed_AU6473(0, 1, 60, "4Bits")
                           
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
                    End If
                End If
                      
                
                Tester.Print "rv0="; rv0
                     
                'If rv0 <> 0 Then
                '    If LightON <> &HBF Or LightOFF <> &HFF Then
                '        Tester.Print "LightON="; LightON
                '        Tester.Print "LightOFF="; LightOFF
                '        UsbSpeedTestResult = GPO_FAIL
                '        rv0 = 3
                '    End If
                'End If
                    
                     
        '=======================================================================================
        'SD R / W
        '=======================================================================================
                Call MsecDelay(0.1)
                
                TmpLBA = LBA
                rv1 = 0
                LBA = LBA + 199
                            
                
                If rv0 = 1 Then
                    rv1 = CBWTest_New_128_Sector_AU6377(1, 1)  ' write
                
                
                    If rv1 <> 1 Then
                        LBA = TmpLBA
                        GoTo AU6371ELResult
                    End If
                End If
                
                LBA = TmpLBA
                ClosePipe
                              
                Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                
        
        '************************************************
        '*               CF Card test                   *
        '************************************************
                   
                rv1 = rv0  'AU6473 no CF slot
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                'Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
        
        '************************************************
        '*               SMC Card test                   *
        '************************************************
              
                rv2 = rv1   'AU6473 no SMC slot
                 
                
              
        '************************************************
        '*               XD Card test                   *
        '************************************************
                
                rv3 = rv2   'AU6473 no XD slot
                
        '************************************************
        '*               MS Card test                   *
        '************************************************
                   
                     
                rv4 = rv3  'AU6473 has no MS slot pin
               
                'Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
        
        '************************************************
        '*               MS PRO Card test               *
        '************************************************
              
                
                
                CardResult = DO_WritePort(card, Channel_P1A, &H7F)
              
                Call MsecDelay(0.2)
               
                CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                 
                Call MsecDelay(0.1)
                'OpenPipe
                '    rv5 = ReInitial(0)
                'ClosePipe
                
                
                    rv5 = CBWTest_New(0, rv4, ChipString)
               
                If rv5 = 1 Then
                    rv5 = Read_MS_Speed_AU6473(0, 0, 60, "4Bits")
                    
                    If rv5 <> 1 Then
                        rv5 = 2
                        Tester.Print "MS bus width Fail"
                    End If
                End If
                
                Call LabelMenu(31, rv5, rv4)
                
                Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                 
               

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
Public Sub AU6473BLF20TestSub()

Dim TmpLBA As Long
Dim i As Integer
      
    Tester.Print "AU6473BL Test Begin ..."

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
                
            TmpLBA = LBA
            LBA = LBA + 1
        
        '=========================================
        '    POWER on
        '=========================================
            
            'CardResult = DO_WritePort(card, Channel_P1A, &HFF)
            'Call MsecDelay(0.2)
           
            If CardResult <> 0 Then
                MsgBox "Power off fail"
                End
            End If
                 
            If GetDeviceName(ChipString) <> "" Then
                rv0 = 0
                Call LabelMenu(0, rv0, 1)
                GoTo AU6371ELResult
            End If
                 
            'Call MsecDelay(0.1)
                 
            'CardResult = DO_WritePort(card, Channel_P1A, &H7E)          'power-on
                  
            'Call MsecDelay(2#)       'power on time
                 
            CardResult = DO_WritePort(card, Channel_P1A, &HCE)          'LUN0_SD + LUN1_SD power-on
                  
            Call MsecDelay(2.6)         'power on time
            
            ChipString = "vid"
            
        '************************************************
        '*               SD Card test                   *
        '************************************************
        
            'Call MsecDelay(0.3)
             
            '===========================================
            'NO card test
            '============================================
  
                ' set Card0_SD card detect down
                'CardResult = DO_WritePort(card, Channel_P1A, &HDE)      'LUN0_SD  power-on
                'Call MsecDelay(0.3)
                'CardResult = DO_WritePort(card, Channel_P1A, &HCE)      'LUN0_SD + LUN1_SD power-on
                'Call MsecDelay(0.3)
                      
                If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect Down Fail"
                    End
                End If
                     
                'Call MsecDelay(1.2)
                     
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                
                If CardResult <> 0 Then
                    MsgBox "Read light On fail"
                    End
                End If
                           
                If LightOn <> 254 Then
                    Tester.Print "LightON="; LightOn
                    UsbSpeedTestResult = GPO_FAIL
                    rv0 = 3
                    GoTo AU6371ELResult
                End If
                           
                'ClosePipe
                
                rv0 = CBWTest_New(0, 1, ChipString)
                
                If rv0 = 1 Then
                    rv0 = Read_SD_Speed_AU6473(TmpLBA, 0, 60, "8Bits")   'Lun0 SD
                    Tester.Print "LUN0 MMC= "; rv0
                    
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
                    ElseIf rv0 = 1 Then
                        
                        rv0 = CBWTest_New_128_Sector_AU6377(0, rv0)  ' write          LUN0_SD
                        
                        If rv0 <> 1 Then
                            rv0 = 2
                            LBA = TmpLBA
                        End If
                    End If
                    
                    ClosePipe
                    Call MsecDelay(0.2)
                    
                    If rv0 = 1 Then
                        rv0 = CBWTest_New(1, rv0, ChipString)
                        
                        If rv0 = 1 Then
                            rv0 = Read_SD_Speed_AU6473(TmpLBA, 1, 60, "4Bits")   'Lun1 SD
                            Tester.Print "LUN1 SD= "; rv0
                        End If
                        
                        If rv0 <> 1 Then
                            rv0 = 2
                            Tester.Print "SD bus width Fail"
                        Else
                            rv0 = CBWTest_New_128_Sector_AU6377(1, rv0)  ' write          LUN1_SD
                            
                            If rv0 <> 1 Then
                                rv0 = 2
                                LBA = TmpLBA
                            End If
                        
                        End If
                    End If
                End If
                      
                ClosePipe
                
                Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                
        
        '************************************************
        '*               CF Card test                   *
        '************************************************
                   
                rv1 = rv0  'AU6473 no CF slot
                 
                ' Call LabelMenu(1, rv1, rv0)
            
                'Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
        
        '************************************************
        '*               SMC Card test                   *
        '************************************************
              
                rv2 = rv1   'AU6473 no SMC slot
                 
                
              
        '************************************************
        '*               XD Card test                   *
        '************************************************
            If rv2 = 1 Then
                
                'CardResult = DO_WritePort(card, Channel_P1A, &HCA)  'close LUN0_SD + LUN1_SD + LUN0_XD
                'Call MsecDelay(0.1)
                
                OpenPipe
                rv3 = ReInitial(0)
                rv3 = ReInitial(1)
                ClosePipe
                
                CardResult = DO_WritePort(card, Channel_P1A, &HFE)  'close LUN0_SD + LUN1_SD
                Call MsecDelay(0.2)
                
                CardResult = DO_WritePort(card, Channel_P1A, &HFA)  'LUN0_XD power-on
                Call MsecDelay(0.3)
                
                
                
                If rv3 = 1 Then
                    rv3 = CBWTest_New(0, rv2, ChipString)
                End If
                
                ClosePipe
                
            End If
                
                Call LabelMenu(0, rv3, rv2)   ' no card test fail
            
                     
                Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                
                
        '************************************************
        '*               MS Card test                   *
        '************************************************
                   
                     
                rv4 = rv3  'AU6473 has no MS slot pin
                
        
        '************************************************
        '*               MS PRO Card test               *
        '************************************************
            If rv4 = 1 Then
                
                'CardResult = DO_WritePort(card, Channel_P1A, &HB2)  'LUN0_XD  +  LUN0_MS + LUN1_MS power-on
                'Call MsecDelay(0.1)
                
                OpenPipe
                rv5 = ReInitial(0)
                ClosePipe
                
                CardResult = DO_WritePort(card, Channel_P1A, &HFE)  'close LUN0_MS + LUN1_MS power-on
                Call MsecDelay(0.2)
                
                CardResult = DO_WritePort(card, Channel_P1A, &HB6)  'LUN0_MS + LUN1_MS power-on
                Call MsecDelay(0.3)
                
                
                
                If rv5 = 1 Then
                    rv5 = CBWTest_New(0, rv4, ChipString)   'LUN0_MSpro
                    
                    If rv5 = 1 Then
                        rv5 = Read_MS_Speed_AU6473(TmpLBA, 0, 60, "4Bits")
                        
                        If rv5 <> 1 Then
                            rv5 = 2
                            Tester.Print "MS bus width Fail"
                            Call LabelMenu(1, rv5, rv4)
                        Else
                            Tester.Print "LUN0 MSpro PASS"
                        End If
                    
                    End If
                End If
                
                ClosePipe
                
                Call MsecDelay(0.2)
                If rv5 = 1 Then
                    rv5 = CBWTest_New(1, rv4, ChipString)   'LUN1_MSpro
                    
                    If rv5 = 1 Then
                        rv5 = Read_MS_Speed_AU6473(TmpLBA, 1, 60, "4Bits")
                         
                        If rv5 <> 1 Then
                            rv5 = 2
                            Tester.Print "MS bus width Fail"
                            Call LabelMenu(1, rv5, rv4)
                        Else
                            Tester.Print "LUN1 MSpro PASS"
                        End If
                    End If
                End If
            End If
                
                ClosePipe
                Call LabelMenu(31, rv5, rv4)
                
                Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                
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
