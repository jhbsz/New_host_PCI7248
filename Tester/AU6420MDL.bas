Attribute VB_Name = "AU6420MDL"
Public Sub AU6420ALTestSub()
   
Dim TmpChip As String
Dim RomSelector As Byte
               
  If ChipName = "AU6370GLF20" Then
      ChipName = "AU6370DLF20"
  End If
                
                ' open power
 If ChipName = "AU6377ALF24" Or ChipName = "AU6377ALF25" Then
     TmpChip = ChipName
     ChipName = "AU6376"
 End If
                
                
            '    PowerSet (1) ' for 3.3V , 2.5 V
 If ChipName = "AU6370DLF20" Or ChipName = "AU6378ALF20" Or ChipName = "AU6378ALF21" Or ChipName = "AU6420ALF20" Then
     TmpChip = ChipName
     ChipName = "AU6376"
 End If
            
                'GPIO control setting
If ChipName = "AU6370BL" Or InStr(ChipName, "AU6375HL") <> 0 Or ChipName = "AU6375CL" Or ChipName = "AU6377ALF21" Or ChipName = "AU6377ALS10" Then
     TmpChip = ChipName
     ChipName = "AU6376"
End If
                
If ChipName = "AU6376ELF20" Or ChipName = "AU6376ILF20" Then
      ChipName = "AU6376"
End If
                
If ChipName = "AU6376JLF20" Then
      TmpChip = ChipName
     ChipName = "AU6376"
End If
                
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    
If ChipName = "AU6368A" Then
       CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 0111 1111
           result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
End If
If ChipName = "AU6368A1" Or ChipName = "AU6376" Then
        result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
    CardResult = DO_WritePort(card, Channel_P1A, &H3E)  ' 1111 1110
End If
                  
 
                  
 If Left(TmpChip, 8) = "AU6378AL" Or Left(TmpChip, 8) = "AU6420AL" Then
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 1111 1110
         CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' 1111 1110
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 1111 1110
  End If
  
  
  '========================== AU6377 new board switch assign ment  ============
If TmpChip = "AU6377ALF21" Then ' this for new board  and internalrom
         RomSelector = &H10  '-------- this is for MS in pin
  End If
  
  
  If TmpChip = "AU6377ALF24" Then ' this for new board  and internalrom
         RomSelector = &H10
  End If
  
  If TmpChip = "AU6377ALF25" Then ' this for new board  and internalrom
         RomSelector = &H0
  End If
         
         
  If Left(TmpChip, 10) = "AU6377ALF2" Then
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector)  ' 1111 1110
         CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' 1111 1110
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector) ' 5th bit is rom selector, High is internal rom
  End If
  
  
 
  
  
  
  
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
                Tester.Print LBA
                If TmpChip = "AU6377ALF25" Or TmpChip = "AU6378ALF21" Then
                  VidName = "vid_1984"
                Else
                 VidName = "vid_058f"
                End If
                
              
                ClosePipe
                 rv0 = CBWTest_New_no_card(0, 1, VidName)
                'Tester.print "a1"
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                rv1 = CBWTest_New_no_card(1, rv0, VidName)
               '  Tester.print "a2"
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
                
                rv2 = CBWTest_New_no_card(2, rv1, VidName)
               '  Tester.print "a3"
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
                 rv3 = CBWTest_New_no_card(3, rv2, VidName)
                ' Tester.print "a4"
                ClosePipe
              Call LabelMenu(3, rv3, rv2)
                
 '================================= Test light off =============================
                
                If Left(TmpChip, 10) = "AU6377ALF2" Then
                
                ' test chip
                      '    CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If LightOff <> 255 Then
                            UsbSpeedTestResult = GPO_FAIL
                            rv0 = 2
                         End If
          
                End If
                
                
                If TmpChip = "AU6378ALF20" Or TmpChip = "AU6420ALF20" Then
                
                ' test chip
                      '    CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If LightOff <> 255 Then
                            UsbSpeedTestResult = GPO_FAIL
                            rv0 = 2
                         End If
          
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
                
                Tester.Print "Test Result"; TestResult
                       
       
                 
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv3, " \\MS :0 Unknow device, 1 pass ,2 card change bit fail"
                 
'====================================== Assing R/W test switch =====================================
                 
                If TestResult = "PASS" Then
                  
                   If ChipName = "AU6368A" Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H64)  ' 0110 0100
                   End If
                   
                   If ChipName = "AU6368A1" Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H20)  ' 0010 0000
                   End If
                   
                   If ChipName = "AU6376" Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H10)  ' 0110 0100
                   End If
                   
                   
                    If TmpChip = "AU6376JLF20" Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                   End If
                   
                    If TmpChip = "AU6378ALF20" Or TmpChip = "AU6420ALF20" Then
                        CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
                        Call MsecDelay(0.5)
                    End If
                    
               
                   
                    If Left(TmpChip, 10) = "AU6377ALF2" Then
                        CardResult = DO_WritePort(card, Channel_P1A, &H4 + RomSelector) ' external rom + SMC excluding
                        Call MsecDelay(0.5)
                    End If
                   
                    
                   Call MsecDelay(0.1)
                 
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
                 rv0 = CBWTest_New(0, 1, VidName)     ' SD slot
                 
                 If rv0 = 1 And TmpChip = "AU6375HLF22" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                End If
                
                   If Left(TmpChip, 10) = "AU6377ALF2" Then
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                      LBA = TmpLBA
                   End If
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 rv1 = CBWTest_New(1, rv0, VidName)    ' CF slot
            
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
                rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
                
                '============= SMC test begin =======================================
               
              
                 
                          '--- for SMC
                
                CardResult = DO_WritePort(card, Channel_P1A, &H18)  ' 0110 0100
                Call MsecDelay(0.5)
                ClosePipe
                rv3 = CBWTest_New(2, rv2, VidName)
                Call LabelMenu(2, rv3, rv2)
                ClosePipe
                
            
                
                
              
                
               '=============== SMC test END ==================================================
               
               rv4 = CBWTest_New(3, rv3, VidName)
               ClosePipe
               Call LabelMenu(3, rv4, rv3)
           
                 
     
                 
             
                 
               
                
             
                         
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If rv4 = 1 And LightOff <> 254 Then
                            UsbSpeedTestResult = GPO_FAIL
                            rv4 = 2
                         End If
          
                  
                 
                 
                    
             
                 
        
             
                
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print "LBA="; LBA
                
AU6377ALFResult:
                        
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
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
                
               
                  
                End If
                CardResult = DO_WritePort(card, Channel_P1A, &H1)
                  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
   End Sub
Public Sub AU6420BLS2ASub()
   
Dim TmpChip As String
Dim RomSelector As Byte
Dim SortingMode As Boolean
               
    TmpChip = ChipName
    ChipName = "AU6376"
    
    

'==================================== Switch assign ==========================================
            
SortingMode = False

Call PowerSet2(0, "0.0", "0.2", 1, "0.0", "0.2", 1)
            
If PCI7248InitFinish = 0 Then
    PCI7248Exist
    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
End If
                
'result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
'CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 1111 1110
         'CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' 1111 1110
         Call MsecDelay(0.3)
                    
         'CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 1111 1110
  
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
                Tester.Print LBA
                 
                VidName = "vid_058f"
                
              
                ClosePipe
                 rv0 = CBWTest_New_no_card(0, 1, VidName)
                'Tester.print "a1"
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                rv1 = CBWTest_New_no_card(1, rv0, VidName)
               '  Tester.print "a2"
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
                
                rv2 = CBWTest_New_no_card(2, rv1, VidName)
               '  Tester.print "a3"
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
                 rv3 = CBWTest_New_no_card(3, rv2, VidName)
                ' Tester.print "a4"
                ClosePipe
                Call LabelMenu(3, rv3, rv2)
                
 '================================= Test light off =============================
                
                
                CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                
                If LightOff <> 255 Then
                    UsbSpeedTestResult = GPO_FAIL
                    rv0 = 2
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
                
                Tester.Print "NoCard Test Result"; TestResult
                       
       
                 
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv3, " \\MS :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print " - - - - - - - - - - - - - - - - - - - - - - - - "

'====================================== Assing R/W test switch =====================================
                 
                If TestResult <> "PASS" Then
                    GoTo AU6377ALFResult
                End If
                
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
                    Call MsecDelay(0.4)
                   
                   
                    
                 '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                 '
                 '  R/W test
                 '
                 '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                 
                
                'initial return value
SortingTest:
                
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
                
                rv0 = CBWTest_New(0, 1, VidName)     ' SD slot
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                
                If Not SortingMode Then
                    rv1 = CBWTest_New(1, rv0, VidName)    ' CF slot
                    Call LabelMenu(1, rv1, rv0)
                    ClosePipe
                Else
                    rv1 = 1
                End If
                
                rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
                
                '============= SMC test begin =======================================
                          '--- for SMC
                
                If Not SortingMode Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H18)  ' 0110 0100
                Else
                    CardResult = DO_WritePort(card, Channel_P1A, &H88)  ' close ENA , use extenal PWR
                End If
                
                Call MsecDelay(0.2)
                ClosePipe
                rv3 = CBWTest_New(2, rv2, VidName)
                Call LabelMenu(2, rv3, rv2)
                ClosePipe
                
               '=============== SMC test END ==================================================
               
               rv4 = CBWTest_New(3, rv3, VidName)
               ClosePipe
               Call LabelMenu(3, rv4, rv3)
           
                         
                CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                
                If rv4 = 1 And LightOff <> 254 Then
                    UsbSpeedTestResult = GPO_FAIL
                    rv4 = 2
                End If
          
                    
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print "LBA="; LBA
                
                If (Not SortingMode) And (rv4 = 1) Then
                    
                    'CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                    'MsecDelay (0.1)
                    fnScsi2usb2K_KillEXE
                    CardResult = DO_WritePort(card, Channel_P1A, &H84)
                    MsecDelay (0.1)
                    Tester.Print "Sorting Test ... 3.24V / 1.75V"
                    Call PowerSet2(0, "3.24", "0.2", 1, "1.75", "0.2", 1)
                    MsecDelay (2#)
                    SortingMode = True
                    GoTo SortingTest
                End If
                
AU6377ALFResult:
                        
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
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
                
               
                  
                    Call PowerSet2(0, "0", "0.2", 1, "0", "0.2", 1)
                    CardResult = DO_WritePort(card, Channel_P1A, &H1)
                    
   End Sub
Public Sub AU6420BLS20Sub()
   
'2011/9/8 sorting for IOI RMA(CF can't format) case
   
Dim TmpChip As String

    TmpChip = ChipName
    'ChipName = "AU6376"
    
    '==================================== Switch assign ==========================================
            
    Call PowerSet2(0, "4.2", "0.2", 1, "4.2", "0.2", 1)
            
    If PCI7248InitFinish = 0 Then
        PCI7248Exist
        result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
    End If
                
    CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 1111 1110
    Call MsecDelay(0.3)
                    
    'CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 1111 1110
  
  
    '======================== Begin test ============================================
                  
    'Call MsecDelay(1)
    rv0 = WaitDevOn("vid_058f")
    Call MsecDelay(0.2)

    LBA = LBA + 1
                
                
    '//////////////////////////////////////////////////
    '
    '   no card insert
    '
    '/////////////////////////////////////////////////
                
    Dim VidName As String
    Tester.Print LBA
                 
    VidName = "vid_058f"
                
    ClosePipe

    rv0 = CBWTest_New_no_card(0, 1, VidName)
    ClosePipe
    Call LabelMenu(0, rv0, 1)

    rv1 = CBWTest_New_no_card(1, rv0, VidName)
    ClosePipe
    Call LabelMenu(1, rv1, rv0)
                
    rv2 = CBWTest_New_no_card(2, rv1, VidName)
    ClosePipe
    Call LabelMenu(2, rv2, rv1)

    rv3 = CBWTest_New_no_card(3, rv2, VidName)
    ClosePipe
    Call LabelMenu(3, rv3, rv2)
                
    '================================= Test light off =============================
                
    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                
    If LightOff <> 255 Then
        UsbSpeedTestResult = GPO_FAIL
        rv0 = 2
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
                
    Tester.Print "NoCard Test Result"; TestResult
                 
    Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 card change bit fail"
    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 card change bit fail"
    Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 card change bit fail"
    Tester.Print rv3, " \\MS :0 Unknow device, 1 pass ,2 card change bit fail"
    Tester.Print " - - - - - - - - - - - - - - - - - - - - - - - - "

    '====================================== Assing R/W test switch =====================================
                 
    If TestResult <> "PASS" Then
        GoTo AU6377ALFResult
    End If
                
    CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
    Call MsecDelay(0.4)
                    
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
    
    rv0 = CBWTest_New(0, 1, VidName)     ' SD slot
    ClosePipe
    Call LabelMenu(0, rv0, 1)
                
    rv1 = CBWTest_New(1, rv0, VidName)    'CF slot
    If rv1 = 1 Then
        rv1 = CBWTest_New_128_Sector_AU6377(1, rv1)
    End If
    ClosePipe
    Call LabelMenu(1, rv1, rv0)
                
    rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
    ClosePipe
    Call LabelMenu(2, rv2, rv1)
                
    '============= SMC test begin =======================================
    '--- for SMC
                
    CardResult = DO_WritePort(card, Channel_P1A, &H18)  ' 0110 0100
    Call MsecDelay(0.2)
    ClosePipe
    rv3 = CBWTest_New(2, rv2, VidName)
    ClosePipe
    Call LabelMenu(2, rv3, rv2)
                
    '=============== SMC test END ==================================================
               
    
    rv4 = CBWTest_New(3, rv3, VidName)
    ClosePipe
    Call LabelMenu(3, rv4, rv3)
           
                         
    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                
    If rv4 = 1 And LightOff <> 254 Then
        UsbSpeedTestResult = GPO_FAIL
        rv4 = 2
    End If
          
                    
    Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    Tester.Print "LBA="; LBA
                
                
AU6377ALFResult:
                        
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
    ElseIf rv4 = WRITE_FAIL Then
        MSWriteFail = MSWriteFail + 1
        TestResult = "MS_WF"
    ElseIf rv4 = READ_FAIL Then
        MSReadFail = MSReadFail + 1
        TestResult = "MS_RF"
    ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
                  
    Call PowerSet2(0, "0", "0.2", 1, "0", "0.2", 1)
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                    
   End Sub

Public Sub AU6420BLS30Sub()
   
'2011/9/8 sorting for IOI RMA(CF can't format) case
   
Dim TmpChip As String

    TmpChip = ChipName
    'ChipName = "AU6376"
    
    '==================================== Switch assign ==========================================
            
    If PCI7248InitFinish = 0 Then
        PCI7248Exist
        result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
    End If
                
    CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 1111 1110
    Call MsecDelay(0.3)
                    
    'CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 1111 1110
  
  
    '======================== Begin test ============================================
                  
    'Call MsecDelay(1)
    rv0 = WaitDevOn("vid_058f")
    Call MsecDelay(0.2)

    LBA = LBA + 1
                
                
    '//////////////////////////////////////////////////
    '
    '   no card insert
    '
    '/////////////////////////////////////////////////
                
    Dim VidName As String
    Tester.Print LBA
                 
    VidName = "vid_058f"
                
    ClosePipe

    rv0 = CBWTest_New_no_card(0, 1, VidName)
    ClosePipe
    Call LabelMenu(0, rv0, 1)

    rv1 = CBWTest_New_no_card(1, rv0, VidName)
    ClosePipe
    Call LabelMenu(1, rv1, rv0)
                
    rv2 = CBWTest_New_no_card(2, rv1, VidName)
    ClosePipe
    Call LabelMenu(2, rv2, rv1)

    rv3 = CBWTest_New_no_card(3, rv2, VidName)
    ClosePipe
    Call LabelMenu(3, rv3, rv2)
                
    '================================= Test light off =============================
                
    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                
    If LightOff <> 255 Then
        UsbSpeedTestResult = GPO_FAIL
        rv0 = 2
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
                
    Tester.Print "NoCard Test Result"; TestResult
                 
    Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 card change bit fail"
    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 card change bit fail"
    Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 card change bit fail"
    Tester.Print rv3, " \\MS :0 Unknow device, 1 pass ,2 card change bit fail"
    Tester.Print " - - - - - - - - - - - - - - - - - - - - - - - - "

    '====================================== Assing R/W test switch =====================================
                 
    If TestResult <> "PASS" Then
        GoTo AU6377ALFResult
    End If
                
    CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
    Call MsecDelay(0.4)
                    
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
    
    rv0 = CBWTest_New(0, 1, VidName)     ' SD slot
    ClosePipe
    Call LabelMenu(0, rv0, 1)
                
    rv1 = CBWTest_New(1, rv0, VidName)    'CF slot
    If rv1 = 1 Then
        Call MsecDelay(2#)
        rv1 = CBWTest_New_128_Sector_AU6377(1, rv1)
    End If
    ClosePipe
    Call LabelMenu(1, rv1, rv0)
                
    rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
    ClosePipe
    Call LabelMenu(2, rv2, rv1)
                
    '============= SMC test begin =======================================
    '--- for SMC
                
    CardResult = DO_WritePort(card, Channel_P1A, &H18)  ' 0110 0100
    Call MsecDelay(0.2)
    ClosePipe
    rv3 = CBWTest_New(2, rv2, VidName)
    ClosePipe
    Call LabelMenu(2, rv3, rv2)
                
    '=============== SMC test END ==================================================
               
    
    rv4 = CBWTest_New(3, rv3, VidName)
    ClosePipe
    Call LabelMenu(3, rv4, rv3)
           
                         
    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                
                
    If rv4 = 1 And LightOff <> 254 Then
        UsbSpeedTestResult = GPO_FAIL
        rv4 = 2
    End If
          
                    
    Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    Tester.Print "LBA="; LBA
                
                
AU6377ALFResult:
                        
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
    ElseIf rv4 = WRITE_FAIL Then
        MSWriteFail = MSWriteFail + 1
        TestResult = "MS_WF"
    ElseIf rv4 = READ_FAIL Then
        MSReadFail = MSReadFail + 1
        TestResult = "MS_RF"
    ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
                  
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                    
   End Sub

Public Sub AU6420BLF2ASub()
   
'2011/9/21 sorting for IOI RMA(CF can't format) case, Loop Read command
   
Dim TmpChip As String
Dim LoopCount As Integer
Const LoopLimit = 50

    TmpChip = ChipName
    'ChipName = "AU6376"
    
    '==================================== Switch assign ==========================================
            
    If PCI7248InitFinish = 0 Then
        PCI7248Exist
        result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
    End If
                
    CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 1111 1110
    Call MsecDelay(0.3)
                    
    'CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 1111 1110
  
  
    '======================== Begin test ============================================
                  
    'Call MsecDelay(1)
    rv0 = WaitDevOn("vid_058f")
    Call MsecDelay(0.2)

    LBA = LBA + 1
                
                
    '//////////////////////////////////////////////////
    '
    '   no card insert
    '
    '/////////////////////////////////////////////////
                
    Dim VidName As String
    Tester.Print LBA
                 
    VidName = "vid_058f"
                
    ClosePipe

    rv0 = CBWTest_New_no_card(0, 1, VidName)
    ClosePipe
    Call LabelMenu(0, rv0, 1)

    rv1 = CBWTest_New_no_card(1, rv0, VidName)
    ClosePipe
    Call LabelMenu(1, rv1, rv0)
                
    rv2 = CBWTest_New_no_card(2, rv1, VidName)
    ClosePipe
    Call LabelMenu(2, rv2, rv1)

    rv3 = CBWTest_New_no_card(3, rv2, VidName)
    ClosePipe
    Call LabelMenu(3, rv3, rv2)
                
    '================================= Test light off =============================
                
    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                
    If LightOff <> 255 Then
        UsbSpeedTestResult = GPO_FAIL
        rv0 = 2
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
                
    Tester.Print "NoCard Test Result"; TestResult
                 
    Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 card change bit fail"
    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 card change bit fail"
    Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 card change bit fail"
    Tester.Print rv3, " \\MS :0 Unknow device, 1 pass ,2 card change bit fail"
    Tester.Print " - - - - - - - - - - - - - - - - - - - - - - - - "

    '====================================== Assing R/W test switch =====================================
                 
    If TestResult <> "PASS" Then
        GoTo AU6377ALFResult
    End If
                
    CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
    Call MsecDelay(0.4)
                    
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
    
    rv0 = CBWTest_New(0, 1, VidName)     ' SD slot
    ClosePipe
    Call LabelMenu(0, rv0, 1)
                
    rv1 = CBWTest_New(1, rv0, VidName)    'CF slot
    If rv1 = 1 Then
        For LoopCount = 0 To LoopLimit
        'Call MsecDelay(2#)
        'rv1 = CBWTest_New_128_Sector_AU6377(1, rv1)
            rv1 = Read_Data_AssignSize(LBA, 1, 2048)
            'LBA = LBA + 200
            Call MsecDelay(0.05)
            If rv1 <> 1 Then
                Tester.Print "Fail Cycle: " & LoopCount
                Exit For
            End If
        Next
    End If
    
    ClosePipe
    Call LabelMenu(1, rv1, rv0)
                
    rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
    ClosePipe
    Call LabelMenu(2, rv2, rv1)
                
    '============= SMC test begin =======================================
    '--- for SMC
                
    CardResult = DO_WritePort(card, Channel_P1A, &H18)  ' 0110 0100
    Call MsecDelay(0.2)
    ClosePipe
    rv3 = CBWTest_New(2, rv2, VidName)
    ClosePipe
    Call LabelMenu(2, rv3, rv2)
                
    '=============== SMC test END ==================================================
               
    
    rv4 = CBWTest_New(3, rv3, VidName)
    ClosePipe
    Call LabelMenu(3, rv4, rv3)
           
                         
    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                
                
    If rv4 = 1 And LightOff <> 254 Then
        UsbSpeedTestResult = GPO_FAIL
        rv4 = 2
    End If
          
                    
    Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    Tester.Print "LBA="; LBA
                
                
AU6377ALFResult:
                        
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
    ElseIf rv4 = WRITE_FAIL Then
        MSWriteFail = MSWriteFail + 1
        TestResult = "MS_WF"
    ElseIf rv4 = READ_FAIL Then
        MSReadFail = MSReadFail + 1
        TestResult = "MS_RF"
    ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
                  
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                    
   End Sub
Public Sub AU6420BLTestSub()
   
Dim TmpChip As String
Dim RomSelector As Byte
               
  If ChipName = "AU6370GLF20" Then
      ChipName = "AU6370DLF20"
  End If
                
                ' open power
 If ChipName = "AU6377ALF24" Or ChipName = "AU6377ALF25" Then
     TmpChip = ChipName
     ChipName = "AU6376"
 End If
                
                
            '    PowerSet (1) ' for 3.3V , 2.5 V
 If ChipName = "AU6370DLF20" Or ChipName = "AU6378ALF20" Or ChipName = "AU6378ALF21" Or ChipName = "AU6420BLF20" Then
     TmpChip = ChipName
     ChipName = "AU6376"
 End If
            
                'GPIO control setting
If ChipName = "AU6370BL" Or InStr(ChipName, "AU6375HL") <> 0 Or ChipName = "AU6375CL" Or ChipName = "AU6377ALF21" Or ChipName = "AU6377ALS10" Then
     TmpChip = ChipName
     ChipName = "AU6376"
End If
                
If ChipName = "AU6376ELF20" Or ChipName = "AU6376ILF20" Then
      ChipName = "AU6376"
End If
                
If ChipName = "AU6376JLF20" Then
      TmpChip = ChipName
     ChipName = "AU6376"
End If
                
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    
If ChipName = "AU6368A" Then
       CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 0111 1111
           result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
End If
If ChipName = "AU6368A1" Or ChipName = "AU6376" Then
        result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
    CardResult = DO_WritePort(card, Channel_P1A, &H3E)  ' 1111 1110
End If
                  
 
                  
 If Left(TmpChip, 8) = "AU6378AL" Or Left(TmpChip, 8) = "AU6420BL" Then
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 1111 1110
         CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' 1111 1110
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 1111 1110
  End If
  
  
  '========================== AU6377 new board switch assign ment  ============
If TmpChip = "AU6377ALF21" Then ' this for new board  and internalrom
         RomSelector = &H10  '-------- this is for MS in pin
  End If
  
  
  If TmpChip = "AU6377ALF24" Then ' this for new board  and internalrom
         RomSelector = &H10
  End If
  
  If TmpChip = "AU6377ALF25" Then ' this for new board  and internalrom
         RomSelector = &H0
  End If
         
         
  If Left(TmpChip, 10) = "AU6377ALF2" Then
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector)  ' 1111 1110
         CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' 1111 1110
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector) ' 5th bit is rom selector, High is internal rom
  End If
  
  
 
  
  
  
  
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
                Tester.Print LBA
                If TmpChip = "AU6377ALF25" Or TmpChip = "AU6378ALF21" Then
                  VidName = "vid_1984"
                Else
                 VidName = "vid_058f"
                End If
                
              
                ClosePipe
                 rv0 = CBWTest_New_no_card(0, 1, VidName)
                'Tester.print "a1"
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                rv1 = CBWTest_New_no_card(1, rv0, VidName)
               '  Tester.print "a2"
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
                
                rv2 = CBWTest_New_no_card(2, rv1, VidName)
               '  Tester.print "a3"
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
                 rv3 = CBWTest_New_no_card(3, rv2, VidName)
                ' Tester.print "a4"
                ClosePipe
              Call LabelMenu(3, rv3, rv2)
                
 '================================= Test light off =============================
                
                If Left(TmpChip, 10) = "AU6377ALF2" Then
                
                ' test chip
                      '    CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If LightOff <> 255 Then
                            UsbSpeedTestResult = GPO_FAIL
                            rv0 = 2
                         End If
          
                End If
                
                
                If TmpChip = "AU6378ALF20" Or TmpChip = "AU6420BLF20" Then
                
                ' test chip
                      '    CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If LightOff <> 255 Then
                            UsbSpeedTestResult = GPO_FAIL
                            rv0 = 2
                         End If
          
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
                
                Tester.Print "Test Result"; TestResult
                       
       
                 
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv3, " \\MS :0 Unknow device, 1 pass ,2 card change bit fail"
                 
'====================================== Assing R/W test switch =====================================
                 
                If TestResult = "PASS" Then
                  
                   If ChipName = "AU6368A" Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H64)  ' 0110 0100
                   End If
                   
                   If ChipName = "AU6368A1" Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H20)  ' 0010 0000
                   End If
                   
                   If ChipName = "AU6376" Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H10)  ' 0110 0100
                   End If
                   
                   
                    If TmpChip = "AU6376JLF20" Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                   End If
                   
                    If TmpChip = "AU6378ALF20" Or TmpChip = "AU6420BLF20" Then
                        CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
                        Call MsecDelay(0.5)
                    End If
                    
               
                   
                    If Left(TmpChip, 10) = "AU6377ALF2" Then
                        CardResult = DO_WritePort(card, Channel_P1A, &H4 + RomSelector) ' external rom + SMC excluding
                        Call MsecDelay(0.5)
                    End If
                   
                    
                   Call MsecDelay(0.1)
                 
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
                 rv0 = CBWTest_New(0, 1, VidName)     ' SD slot
                 
                 If rv0 = 1 And TmpChip = "AU6375HLF22" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                End If
                
                   If Left(TmpChip, 10) = "AU6377ALF2" Then
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                      LBA = TmpLBA
                   End If
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 rv1 = CBWTest_New(1, rv0, VidName)    ' CF slot
            
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
                rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
                
                '============= SMC test begin =======================================
               
              
                 
                          '--- for SMC
                
                CardResult = DO_WritePort(card, Channel_P1A, &H18)  ' 0110 0100
                Call MsecDelay(0.5)
                ClosePipe
                rv3 = CBWTest_New(2, rv2, VidName)
                Call LabelMenu(2, rv3, rv2)
                ClosePipe
                
            
                
                
              
                
               '=============== SMC test END ==================================================
               
               rv4 = CBWTest_New(3, rv3, VidName)
               ClosePipe
               Call LabelMenu(3, rv4, rv3)
           
                 
     
                 
             
                 
               
                
             
                         
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If rv4 = 1 And LightOff <> 254 Then
                            UsbSpeedTestResult = GPO_FAIL
                            rv4 = 2
                         End If
          
                  
                 
                 
                    
             
                 
        
             
                
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print "LBA="; LBA
                
AU6377ALFResult:
                        
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
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
                
               
                  
                End If
                CardResult = DO_WritePort(card, Channel_P1A, &H1)
                  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
   End Sub

Public Sub AU6420CLTestSub()
   
Dim TmpChip As String
Dim RomSelector As Byte
               
  If ChipName = "AU6370GLF20" Then
      ChipName = "AU6370DLF20"
  End If
                
                ' open power
 If ChipName = "AU6377ALF24" Or ChipName = "AU6377ALF25" Then
     TmpChip = ChipName
     ChipName = "AU6376"
 End If
                
                
            '    PowerSet (1) ' for 3.3V , 2.5 V
 If ChipName = "AU6370DLF20" Or ChipName = "AU6378ALF20" Or ChipName = "AU6378ALF21" Or ChipName = "AU6420CLF20" Then
     TmpChip = ChipName
     ChipName = "AU6376"
 End If
            
                'GPIO control setting
If ChipName = "AU6370BL" Or InStr(ChipName, "AU6375HL") <> 0 Or ChipName = "AU6375CL" Or ChipName = "AU6377ALF21" Or ChipName = "AU6377ALS10" Then
     TmpChip = ChipName
     ChipName = "AU6376"
End If
                
If ChipName = "AU6376ELF20" Or ChipName = "AU6376ILF20" Then
      ChipName = "AU6376"
End If
                
If ChipName = "AU6376JLF20" Then
      TmpChip = ChipName
     ChipName = "AU6376"
End If
                
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    
If ChipName = "AU6368A" Then
       CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 0111 1111
           result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
End If
If ChipName = "AU6368A1" Or ChipName = "AU6376" Then
        result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
    CardResult = DO_WritePort(card, Channel_P1A, &H3E)  ' 1111 1110
End If
                  
 
                  
 If Left(TmpChip, 8) = "AU6378AL" Or Left(TmpChip, 8) = "AU6420CL" Then
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 1111 1110
         CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' 1111 1110
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 1111 1110
  End If
  
  
  '========================== AU6377 new board switch assign ment  ============
If TmpChip = "AU6377ALF21" Then ' this for new board  and internalrom
         RomSelector = &H10  '-------- this is for MS in pin
  End If
  
  
  If TmpChip = "AU6377ALF24" Then ' this for new board  and internalrom
         RomSelector = &H10
  End If
  
  If TmpChip = "AU6377ALF25" Then ' this for new board  and internalrom
         RomSelector = &H0
  End If
         
         
  If Left(TmpChip, 10) = "AU6377ALF2" Then
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector)  ' 1111 1110
         CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' 1111 1110
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector) ' 5th bit is rom selector, High is internal rom
  End If
  
  


  
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
                Tester.Print LBA
                If TmpChip = "AU6377ALF25" Or TmpChip = "AU6378ALF21" Then
                  VidName = "vid_1984"
                Else
                 VidName = "vid_058f"
                End If
                
              
                ClosePipe
                 rv0 = CBWTest_New_no_card(0, 1, VidName)
                'Tester.print "a1"
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                rv1 = CBWTest_New_no_card(1, rv0, VidName)
               '  Tester.print "a2"
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
               
               rv2 = CBWTest_New_no_card(2, rv1, VidName)
               '  Tester.print "a3"
                 Call LabelMenu(2, rv2, rv1)
                ClosePipe
                  rv3 = CBWTest_New_no_card(3, rv2, VidName)
                ' Tester.print "a4"
                ClosePipe
               Call LabelMenu(3, rv3, rv2)
                
 '================================= Test light off =============================
                
                If Left(TmpChip, 10) = "AU6377ALF2" Then
                
                ' test chip
                      '    CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If LightOff <> 255 Then
                            UsbSpeedTestResult = GPO_FAIL
                            rv0 = 2
                         End If
          
                End If
                
                
                If TmpChip = "AU6378ALF20" Or TmpChip = "AU6420CLF20" Then
                
                ' test chip
                      '    CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If LightOff <> 255 Then
                            UsbSpeedTestResult = GPO_FAIL
                            rv0 = 2
                         End If
          
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
                
                Tester.Print "Test Result"; TestResult
                       
       
                 
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv3, " \\MS :0 Unknow device, 1 pass ,2 card change bit fail"
                 
'====================================== Assing R/W test switch =====================================
                 
                If TestResult = "PASS" Then
                  
                   If ChipName = "AU6368A" Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H64)  ' 0110 0100
                   End If
                   
                   If ChipName = "AU6368A1" Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H20)  ' 0010 0000
                   End If
                   
                   If ChipName = "AU6376" Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H10)  ' 0110 0100
                   End If
                   
                   
                    If TmpChip = "AU6376JLF20" Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                   End If
                   
                    If TmpChip = "AU6378ALF20" Or TmpChip = "AU6420CLF20" Then
                        CardResult = DO_WritePort(card, Channel_P1A, &H7C)  ' 0110 0100
                        Call MsecDelay(0.5)
                    End If
                    
               
                   
                    If Left(TmpChip, 10) = "AU6377ALF2" Then
                        CardResult = DO_WritePort(card, Channel_P1A, &H4 + RomSelector) ' external rom + SMC excluding
                        Call MsecDelay(0.5)
                    End If
                   
                    
                   Call MsecDelay(0.1)
                 
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
                 rv0 = CBWTest_New(0, 1, VidName)     ' cf slot
                 
                 If rv0 = 1 And TmpChip = "AU6375HLF22" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                End If
                
                   If Left(TmpChip, 10) = "AU6377ALF2" Then
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                      LBA = TmpLBA
                   End If
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 rv1 = CBWTest_New(1, rv0, VidName)    ' sd slot
            
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
                  CardResult = DO_WritePort(card, Channel_P1A, &H7D)  ' 0110 0100
                Call MsecDelay(0.1)
                   CardResult = DO_WritePort(card, Channel_P1A, &H75)  ' 0110 0100
                Call MsecDelay(0.1)
              
                rv2 = CBWTest_New(1, rv1, VidName)    'XD slot
                Call LabelMenu(1, rv2, rv1)
                ClosePipe
                
                '============= SMC test begin =======================================
               
              
                 
                          '--- for SMC
                CardResult = DO_WritePort(card, Channel_P1A, &H7D)  ' 0110 0100
                Call MsecDelay(0.1)
                      
                
                CardResult = DO_WritePort(card, Channel_P1A, &H79)  ' 0110 0100
                Call MsecDelay(0.1)
                ClosePipe
                rv3 = CBWTest_New(1, rv2, VidName)
                Call LabelMenu(2, rv3, rv2)
                ClosePipe
                
            
                
                
              
                
               '=============== SMC test END ==================================================
               
                  CardResult = DO_WritePort(card, Channel_P1A, &H7D)  ' 0110 0100
                Call MsecDelay(0.1)
                      
                
                CardResult = DO_WritePort(card, Channel_P1A, &H5D)  ' 0110 0100
                Call MsecDelay(0.1)
               
               rv4 = CBWTest_New(1, rv3, VidName)
               ClosePipe
               Call LabelMenu(3, rv4, rv3)
           
                 
     
         
                         
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If rv4 = 1 And LightOff <> 254 Then
                            UsbSpeedTestResult = GPO_FAIL
                            rv4 = 2
                         End If
          
                  
                 
                 
                    
             
                 
        
             
                
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print "LBA="; LBA
                
AU6377ALFResult:
                        
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
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
                
               
                  
                End If
                CardResult = DO_WritePort(card, Channel_P1A, &H1)
                  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
   End Sub

Public Sub AU6476BLTestSub()
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
               
   
TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
  
                  
 If Left(TmpChip, 8) = "AU6378AL" Or Left(TmpChip, 8) = "AU6476BL" Then
 
         CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  // force into reader mode
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.2)
         CardResult = DO_WritePort(card, Channel_P1A, &H0)
         
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H3F)
          
          
  End If
  
  
  '========================== AU6377 new board switch assign ment  ============
If TmpChip = "AU6377ALF21" Then ' this for new board  and internalrom
         RomSelector = &H10  '-------- this is for MS in pin
  End If
  
  
  If TmpChip = "AU6377ALF24" Then ' this for new board  and internalrom
         RomSelector = &H10
  End If
  
  If TmpChip = "AU6377ALF25" Then ' this for new board  and internalrom
         RomSelector = &H0
  End If
         
         
  If Left(TmpChip, 10) = "AU6377ALF2" Then
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector)  ' 1111 1110
         CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' 1111 1110
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector) ' 5th bit is rom selector, High is internal rom
  End If
  
  


  
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1.3)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
                Tester.Print LBA
                If TmpChip = "AU6377ALF25" Or TmpChip = "AU6378ALF21" Then
                  VidName = "vid_1984"
                Else
                 VidName = "vid_058f"
                End If
                
              
                ClosePipe
                 rv0 = CBWTest_New_no_card(0, 1, VidName)
                'Tester.print "a1"
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                rv1 = CBWTest_New_no_card(1, rv0, VidName)
               '  Tester.print "a2"
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
               
               rv2 = CBWTest_New_no_card(2, rv1, VidName)
               '  Tester.print "a3"
                 Call LabelMenu(2, rv2, rv1)
                ClosePipe
                  rv3 = CBWTest_New_no_card(3, rv2, VidName)
                ' Tester.print "a4"
                ClosePipe
               Call LabelMenu(3, rv3, rv2)
                
 '================================= Test light off =============================
                
              
                
                
                If TmpChip = "AU6476BLF20" Then
                
                ' test chip
                      '    CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If LightOff <> 255 Then
                            UsbSpeedTestResult = GPO_FAIL
                             rv0 = 2
                         End If
          
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
               
                       
       
                 
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv3, " \\MS :0 Unknow device, 1 pass ,2 card change bit fail"
                 
'====================================== Assing R/W test switch =====================================
                 
                If TestResult = "PASS" Then
                   TestResult = ""
                   If ChipName = "AU6476BLF20" Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                   End If
              
                    
                   Call MsecDelay(0.1)
                 
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
                 rv0 = CBWTest_New(0, 1, VidName)     ' sd slot
                 
                 If rv0 = 1 And TmpChip = "AU6375HLF22" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                End If
                
                   If Left(TmpChip, 10) = "AU6377ALF2" Then
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                      LBA = TmpLBA
                   End If
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
                
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
              
                
                 rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
               
                ClosePipe
                
                 If rv1 = 1 And rv2 <> 1 Then
                 Tester.Label9.Caption = "SMC FAIL"
                End If
                '============= SMC test begin =======================================
               
              
                 
                          '--- for SMC
                CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
                Call MsecDelay(0.1)
                      
                
                CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
                Call MsecDelay(0.1)
                ClosePipe
                
                
               
                  rv3 = CBWTest_New(2, rv2, VidName)
               
                Call LabelMenu(2, rv3, rv2)
                ClosePipe
                  If rv2 = 1 And rv3 <> 1 Then
                 Tester.Label9.Caption = "XD FAIL"
                End If
                
                
              
                
               '=============== SMC test END ==================================================
               
              
               rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
            
               ClosePipe
               Call LabelMenu(3, rv4, rv3)
               
 '======================== light test ======================
               
               
               
               
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
      If rv4 = 1 Then
          CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off  for HID mode
          result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
          Call MsecDelay(0.4)
          
          CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
                            
    
          Call MsecDelay(1.2)
          CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
          If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
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
                         GoTo AU6377ALFResult
                     End If
                     
                     '=======================
                     '  key press test, it will return 8 when key up, GPI 6 must do low go hi action
                     '========================
                     
                
                     Do
                          CardResult = DO_WritePort(card, Channel_P1A, &H40) 'GPI6 : bit 6: pull high
                          Sleep (200)
                         CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' GPI6 : bit 6: pull low
                         Sleep (1000)
                       
                         ReturnValue = fnInquiryBtnStatus(DeviceHandle)
                         Tester.Print i; Space(5); "Key press value="; ReturnValue
                         i = i + 1
                     Loop While i < 3 And ReturnValue <> 12
                    ' fnFreeDeviceHandle (DeviceHandle)
                   '  fnFreeDeviceHandle (DeviceHandle)
                     
                     If ReturnValue <> 12 Then
                     
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
                
AU6377ALFResult:
                        
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
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        
                        End If
                
               
                  
                End If
                
                
                
     '    CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off
     '    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
     '    Call MsecDelay(0.8)
     '    CardResult = DO_WritePort(card, Channel_P1A, &H40)
         
                
                
               
   End Sub
   
   Public Sub AU6476BLF21TestSub()
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
               
   
TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
  
                  
 If Left(TmpChip, 8) = "AU6378AL" Or Left(TmpChip, 8) = "AU6476BL" Then
 
         CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  // force into reader mode
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.2)
         CardResult = DO_WritePort(card, Channel_P1A, &H0)
         
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H3F)
          
          
  End If
  
  
  '========================== AU6377 new board switch assign ment  ============
If TmpChip = "AU6377ALF21" Then ' this for new board  and internalrom
         RomSelector = &H10  '-------- this is for MS in pin
  End If
  
  
  If TmpChip = "AU6377ALF24" Then ' this for new board  and internalrom
         RomSelector = &H10
  End If
  
  If TmpChip = "AU6377ALF25" Then ' this for new board  and internalrom
         RomSelector = &H0
  End If
         
         
  If Left(TmpChip, 10) = "AU6377ALF2" Then
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector)  ' 1111 1110
         CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' 1111 1110
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector) ' 5th bit is rom selector, High is internal rom
  End If
  
  


  
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1.3)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
                Tester.Print LBA
                If TmpChip = "AU6377ALF25" Or TmpChip = "AU6378ALF21" Then
                  VidName = "vid_1984"
                Else
                 VidName = "vid_058f"
                End If
                
              
                ClosePipe
                 rv0 = CBWTest_New_no_card(0, 1, VidName)
                'Tester.print "a1"
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                rv1 = CBWTest_New_no_card(1, rv0, VidName)
               '  Tester.print "a2"
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
               
               rv2 = CBWTest_New_no_card(2, rv1, VidName)
               '  Tester.print "a3"
                 Call LabelMenu(2, rv2, rv1)
                ClosePipe
                  rv3 = CBWTest_New_no_card(3, rv2, VidName)
                ' Tester.print "a4"
                ClosePipe
               Call LabelMenu(3, rv3, rv2)
                
 '================================= Test light off =============================
                
              
                
                
                If TmpChip = "AU6476BLF20" Then
                
                ' test chip
                      '    CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If LightOff <> 255 Then
                            UsbSpeedTestResult = GPO_FAIL
                             rv0 = 2
                         End If
          
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
               
                       
       
                 
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv3, " \\MS :0 Unknow device, 1 pass ,2 card change bit fail"
                 
'====================================== Assing R/W test switch =====================================
                 
                If TestResult = "PASS" Then
                   TestResult = ""
                   
                    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                  
                    
                   Call MsecDelay(0.1)
                 
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
                 rv0 = CBWTest_New(0, 1, VidName)     ' sd slot
                 
                 If rv0 = 1 And TmpChip = "AU6375HLF22" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                End If
                
                   If Left(TmpChip, 10) = "AU6377ALF2" Then
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                      LBA = TmpLBA
                   End If
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
                
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
              
                
                 rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
               
                ClosePipe
                
                 If rv1 = 1 And rv2 <> 1 Then
                 Tester.Label9.Caption = "SMC FAIL"
                End If
                '============= SMC test begin =======================================
               
              
                 
                          '--- for SMC
                CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
                Call MsecDelay(0.1)
                      
                
                CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
                Call MsecDelay(0.1)
                ClosePipe
                
                
               
                  rv3 = CBWTest_New(2, rv2, VidName)
               
                Call LabelMenu(2, rv3, rv2)
                ClosePipe
                  If rv2 = 1 And rv3 <> 1 Then
                 Tester.Label9.Caption = "XD FAIL"
                End If
                
                
              
                
               '=============== SMC test END ==================================================
               
              
               rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
            
               ClosePipe
               Call LabelMenu(3, rv4, rv3)
               
 '======================== light test ======================
               
               
               
               
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
      If rv4 = 1 Then
          CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off  for HID mode
          result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
          Call MsecDelay(0.4)
          
          CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
                            
    
          Call MsecDelay(1.2)
          CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
          If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
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
                         GoTo AU6377ALFResult
                     End If
                     
                     '=======================
                     '  key press test, it will return 8 when key up, GPI 6 must do low go hi action
                     '========================
                     
                
                     Do
                          CardResult = DO_WritePort(card, Channel_P1A, &H40) 'GPI6 : bit 6: pull high
                          Sleep (200)
                         CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' GPI6 : bit 6: pull low
                         Sleep (1000)
                       
                         ReturnValue = fnInquiryBtnStatus(DeviceHandle)
                         Tester.Print i; Space(5); "Key press value="; ReturnValue
                         i = i + 1
                     Loop While i < 3 And ReturnValue <> 10
                    ' fnFreeDeviceHandle (DeviceHandle)
                   '  fnFreeDeviceHandle (DeviceHandle)
                     
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
                
AU6377ALFResult:
                        
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
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        
                        End If
                
               
                  
                End If
                
                
                
     '    CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off
     '    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
     '    Call MsecDelay(0.8)
     '    CardResult = DO_WritePort(card, Channel_P1A, &H40)
         
                
                
               
   End Sub
   Public Sub AU6476BLF23TestSub()
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
               
   
TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
Tester.Print "AU6476BL is NB mode"
                  
 If Left(TmpChip, 8) = "AU6378AL" Or Left(TmpChip, 8) = "AU6476BL" Then
 
         CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  // force into reader mode
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.2)
         CardResult = DO_WritePort(card, Channel_P1A, &H0)
         
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H3F)
          
          
  End If
  
              
  
  
  '========================== AU6377 new board switch assign ment  ============
If TmpChip = "AU6377ALF21" Then ' this for new board  and internalrom
         RomSelector = &H10  '-------- this is for MS in pin
  End If
  
  
  If TmpChip = "AU6377ALF24" Then ' this for new board  and internalrom
         RomSelector = &H10
  End If
  
  If TmpChip = "AU6377ALF25" Then ' this for new board  and internalrom
         RomSelector = &H0
  End If
         
         
  If Left(TmpChip, 10) = "AU6377ALF2" Then
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector)  ' 1111 1110
         CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' 1111 1110
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector) ' 5th bit is rom selector, High is internal rom
  End If
  
  


  
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1.3)
               
                
                 If GetDeviceName("vid") <> "" Then
                    rv0 = 0
                    Tester.Print "NB mode test Fail"
                    GoTo AU6377ALFResult
                  
                  End If
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
                Tester.Print LBA
                If TmpChip = "AU6377ALF25" Or TmpChip = "AU6378ALF21" Then
                  VidName = "vid_1984"
                Else
                 VidName = "vid_058f"
                End If
                
              
             
 '================================= Test light off =============================
                
              
                
                
                   
                    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                  
                    
                   Call MsecDelay(1.3)
                 
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
                 rv0 = CBWTest_New_SD_Speed(0, 1, VidName, "8Bits")    ' sd slot
                 
                 If rv0 = 1 And TmpChip = "AU6375HLF22" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                End If
                
                   If Left(TmpChip, 10) = "AU6377ALF2" Then
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                      LBA = TmpLBA
                   End If
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                  rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
                  ' rv1 = 1
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
              
             
                 rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
               
                ClosePipe
                
                 If rv1 = 1 And rv2 <> 1 Then
                 Tester.Label9.Caption = "SMC FAIL"
                End If
                '============= SMC test begin =======================================
               
              
                 
                          '--- for SMC
                CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
                Call MsecDelay(0.1)
                      
                
                CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
                Call MsecDelay(0.1)
                ClosePipe
                
                
               
                  rv3 = CBWTest_New(2, rv2, VidName)
               
                Call LabelMenu(2, rv3, rv2)
                ClosePipe
                  If rv2 = 1 And rv3 <> 1 Then
                 Tester.Label9.Caption = "XD FAIL"
                End If
                
                
              
                
               '=============== SMC test END ==================================================
               
              
               rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
            
               ClosePipe
               Call LabelMenu(3, rv4, rv3)
               
 '======================== light test ======================
               
               
               
               
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
      If rv4 = 1 Then
          CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off  for HID mode
          result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
          Call MsecDelay(0.4)
          
          CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
                            
    
          Call MsecDelay(1.2)
          CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
          If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
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
                         GoTo AU6377ALFResult
                     End If
                     
                     '=======================
                     '  key press test, it will return 8 when key up, GPI 6 must do low go hi action
                     '========================
                     
                
                     Do
                          CardResult = DO_WritePort(card, Channel_P1A, &H40) 'GPI6 : bit 6: pull high
                          Sleep (200)
                         CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' GPI6 : bit 6: pull low
                         Sleep (1000)
                       
                         ReturnValue = fnInquiryBtnStatus(DeviceHandle)
                         Tester.Print i; Space(5); "Key press value="; ReturnValue
                         i = i + 1
                     Loop While i < 3 And ReturnValue <> 10
                    ' fnFreeDeviceHandle (DeviceHandle)
                   '  fnFreeDeviceHandle (DeviceHandle)
                     
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
                
AU6377ALFResult:
                        
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
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        
                        End If
                
               
                  
              
                
                
                
     '    CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off
     '    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
     '    Call MsecDelay(0.8)
     '    CardResult = DO_WritePort(card, Channel_P1A, &H40)
         
                
                
               
   End Sub
   
 Public Sub AU6476BLF24TestSub()
'AU6476BLF24 add MS bus width test, and modify SD bus width
 
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
               
   
TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
Tester.Print "AU6476BL is NB mode"
                  
 If Left(TmpChip, 8) = "AU6378AL" Or Left(TmpChip, 8) = "AU6476BL" Then
 
         CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  // force into reader mode
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.2)
         CardResult = DO_WritePort(card, Channel_P1A, &H0)
         
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H3F)
          
          
  End If
  
              
  
  
  '========================== AU6377 new board switch assign ment  ============
If TmpChip = "AU6377ALF21" Then ' this for new board  and internalrom
         RomSelector = &H10  '-------- this is for MS in pin
  End If
  
  
  If TmpChip = "AU6377ALF24" Then ' this for new board  and internalrom
         RomSelector = &H10
  End If
  
  If TmpChip = "AU6377ALF25" Then ' this for new board  and internalrom
         RomSelector = &H0
  End If
         
         
  If Left(TmpChip, 10) = "AU6377ALF2" Then
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector)  ' 1111 1110
         CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' 1111 1110
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector) ' 5th bit is rom selector, High is internal rom
  End If
  
  


  
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1.3)
               
                
                 If GetDeviceName("vid") <> "" Then
                    rv0 = 0
                    Tester.Print "NB mode test Fail"
                    GoTo AU6377ALFResult
                  
                  End If
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
                Tester.Print LBA
                If TmpChip = "AU6377ALF25" Or TmpChip = "AU6378ALF21" Then
                  VidName = "vid_1984"
                Else
                 VidName = "vid_058f"
                End If
                
              
              
              
              
              
              
             
 '================================= Test light off =============================
                
              
                
                
                   
                    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                  
                    
                   Call MsecDelay(1.3)
                 
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
                 rv0 = CBWTest_New(0, 1, VidName)    ' sd slot
                 
                 
                 If rv0 = 1 Then
                    rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                    If rv0 <> 1 Then
                    Tester.Print "SD bus width Fail"
                    rv0 = 2
                    End If
                 End If
                 
                 
                 
                 If rv0 = 1 And TmpChip = "AU6375HLF22" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                End If
                
                   If Left(TmpChip, 10) = "AU6377ALF2" Then
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                      LBA = TmpLBA
                   End If
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                  rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
                  ' rv1 = 1
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
              
             
                 rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
               
                ClosePipe
                
                 If rv1 = 1 And rv2 <> 1 Then
                 Tester.Label9.Caption = "SMC FAIL"
                End If
                '============= SMC test begin =======================================
               
              
                 
                          '--- for SMC
                CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
                Call MsecDelay(0.1)
                      
                
                CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
                Call MsecDelay(0.1)
                ClosePipe
                
                
               
                  rv3 = CBWTest_New(2, rv2, VidName)
               
                Call LabelMenu(2, rv3, rv2)
                ClosePipe
                  If rv2 = 1 And rv3 <> 1 Then
                 Tester.Label9.Caption = "XD FAIL"
                End If
                
                
              
                
               '=============== SMC test END ==================================================
               
              
               rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
               
               If rv4 = 1 Then
               rv4 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                 If rv4 <> 1 Then
                   Tester.Print "MS bus width Fail"
                   rv4 = 2
                 End If
              End If
            
               ClosePipe
               Call LabelMenu(3, rv4, rv3)
               
 '======================== light test ======================
               
               
               
               
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
      If rv4 = 1 Then
          CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off  for HID mode
          result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
          Call MsecDelay(0.4)
          
          CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
                            
    
          Call MsecDelay(1.2)
          CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
          If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
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
                         GoTo AU6377ALFResult
                     End If
                     
                     '=======================
                     '  key press test, it will return 8 when key up, GPI 6 must do low go hi action
                     '========================
                     
                
                     Do
                          CardResult = DO_WritePort(card, Channel_P1A, &H40) 'GPI6 : bit 6: pull high
                          Sleep (200)
                         CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' GPI6 : bit 6: pull low
                         Sleep (1000)
                       
                         ReturnValue = fnInquiryBtnStatus(DeviceHandle)
                         Tester.Print i; Space(5); "Key press value="; ReturnValue
                         i = i + 1
                     Loop While i < 3 And ReturnValue <> 10
                    ' fnFreeDeviceHandle (DeviceHandle)
                   '  fnFreeDeviceHandle (DeviceHandle)
                     
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
                
AU6377ALFResult:
                        
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
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        
                        End If
                
               
                  
              
                
                
                
     '    CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off
     '    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
     '    Call MsecDelay(0.8)
     '    CardResult = DO_WritePort(card, Channel_P1A, &H40)
         
                
                
               
   End Sub
 Public Sub AU6476BLF25TestSub()
'AU6476BLF24 add MS bus width test, and modify SD bus width
'2011/7/4 AU6476BLF25 enhance XD 64k pattern R/W test

On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
               
   
TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
Tester.Print "AU6476BL is NB mode"
                  
 'If Left(TmpChip, 8) = "AU6378AL" Or Left(TmpChip, 8) = "AU6476BL" Then
 
 '        CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  // force into reader mode
 '        result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
 '        Call MsecDelay(0.2)
 '        CardResult = DO_WritePort(card, Channel_P1A, &H0)
 '
 '        Call MsecDelay(0.3)
 '
         CardResult = DO_WritePort(card, Channel_P1A, &H3F)
          
          
  'End If
    
  
'======================== Begin test ============================================
                  
    Call MsecDelay(1.3)
               
                
    If GetDeviceName("vid") <> "" Then
        rv0 = 0
        Tester.Print "NB mode test Fail"
        GoTo AU6377ALFResult
    End If
               
    LBA = LBA + 1
                
                
    '//////////////////////////////////////////////////
    '
    '   no card insert
    '
    '/////////////////////////////////////////////////
                
    Dim VidName As String
    Tester.Print LBA
    VidName = "vid_058f"
              
             
 '================================= Test light off =============================
                   
    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
    Call MsecDelay(1.3)
                 
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
    rv0 = CBWTest_New(0, 1, VidName)    ' sd slot
                 
    If rv0 = 1 Then
        rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
        
        If rv0 <> 1 Then
            Tester.Print "SD bus width Fail"
            rv0 = 2
        End If
    End If
                 
    If rv0 = 1 And TmpChip = "AU6375HLF22" Then
        ClosePipe
        rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
        ClosePipe
    End If
                
    If Left(TmpChip, 10) = "AU6377ALF2" Then
        TmpLBA = LBA
        LBA = 99
        
        For i = 1 To 30
            rv1 = 0
            LBA = LBA + 199
            
            ClosePipe
            rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
            
            If rv1 <> 1 Then
                LBA = TmpLBA
                GoTo AU6377ALFResult
            End If
        Next
            
        LBA = TmpLBA
    
    End If
                
    Call LabelMenu(0, rv0, 1)
    ClosePipe
    
    rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
    ' rv1 = 1
    Call LabelMenu(1, rv1, rv0)
    ClosePipe
             
    rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
    If rv2 = 1 Then
        rv2 = CBWTest_New_128_Sector_AU6377(2, 1)  ' write
    End If
    
    Call LabelMenu(2, rv2, rv1)
               
    ClosePipe
    If rv1 = 1 And rv2 <> 1 Then
        Tester.Label9.Caption = "SMC FAIL"
    End If
    
'============= SMC test begin =======================================
                 
    '--- for SMC
    CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
    Call MsecDelay(0.1)
                
    CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
    Call MsecDelay(0.1)
    ClosePipe
               
    rv3 = CBWTest_New(2, rv2, VidName)
    Call LabelMenu(2, rv3, rv2)
    ClosePipe
    
    If rv2 = 1 And rv3 <> 1 Then
        Tester.Label9.Caption = "XD FAIL"
    End If
                
'=============== SMC test END ==================================================
               
              
    rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
               
    If rv4 = 1 Then
        rv4 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
        If rv4 <> 1 Then
            Tester.Print "MS bus width Fail"
            rv4 = 2
        End If
    End If
            
    ClosePipe
    Call LabelMenu(3, rv4, rv3)
               
 '======================== light test ======================
               
               
               
               
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
    
    If rv4 = 1 Then
        CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  for HID mode
        'result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
        Call MsecDelay(0.4)
          
        CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
        
        Call MsecDelay(1.2)
        CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
        
        If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
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
            GoTo AU6377ALFResult
        End If
                     
        '=======================
        '  key press test, it will return 8 when key up, GPI 6 must do low go hi action
        '========================
                
        Do
            CardResult = DO_WritePort(card, Channel_P1A, &H40) 'GPI6 : bit 6: pull high
            Sleep (200)
            CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' GPI6 : bit 6: pull low
            Sleep (1000)
                       
            ReturnValue = fnInquiryBtnStatus(DeviceHandle)
            Tester.Print i; Space(5); "Key press value="; ReturnValue
            i = i + 1
        Loop While i < 3 And ReturnValue <> 10
            ' fnFreeDeviceHandle (DeviceHandle)
            'fnFreeDeviceHandle (DeviceHandle)
                     
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
                
                
AU6377ALFResult:
                        
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
    ElseIf rv4 = WRITE_FAIL Then
        MSWriteFail = MSWriteFail + 1
        TestResult = "MS_WF"
    ElseIf rv4 = READ_FAIL Then
        MSReadFail = MSReadFail + 1
        TestResult = "MS_RF"
    ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
                           
End Sub

 Public Sub AU6476BLF26TestSub()
'AU6476BLF24 add MS bus width test, and modify SD bus width
'2011/7/4 AU6476BLF25 enhance XD 64k pattern R/W test

'On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
Dim i As Long
Dim j As Long
   
TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
Tester.Print "AU6476BL is NB mode"
                  
 'If Left(TmpChip, 8) = "AU6378AL" Or Left(TmpChip, 8) = "AU6476BL" Then
 
 '        CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  // force into reader mode
 '        result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
 '        Call MsecDelay(0.2)
 '        CardResult = DO_WritePort(card, Channel_P1A, &H0)
 '
 '        Call MsecDelay(0.3)
 '
         
          
          
  'End If
    
  
'======================== Begin test ============================================
    
               
    LBA = LBA + 1
                
                
    '//////////////////////////////////////////////////
    '
    '   no card insert
    '
    '/////////////////////////////////////////////////
                
    Dim VidName As String
    Tester.Print LBA
    VidName = "vid_058f"
              
             
 '================================= Test light off =============================
                   
    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
    WaitDevOn (VidName)
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
    rv0 = CBWTest_New(0, 1, VidName)    ' sd slot
                 
    If rv0 = 1 Then
        rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
        
        If rv0 <> 1 Then
            Tester.Print "SD bus width Fail"
            rv0 = 2
        End If
    End If
                 
    If rv0 = 1 And TmpChip = "AU6375HLF22" Then
        ClosePipe
        rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
        ClosePipe
    End If
                
    If Left(TmpChip, 10) = "AU6377ALF2" Then
        TmpLBA = LBA
        LBA = 99
        
        For i = 1 To 30
            rv1 = 0
            LBA = LBA + 199
            
            ClosePipe
            rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
            
            If rv1 <> 1 Then
                LBA = TmpLBA
                GoTo AU6377ALFResult
            End If
        Next
            
        LBA = TmpLBA
    
    End If
                
    Call LabelMenu(0, rv0, 1)
    ClosePipe
    
    rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
    ' rv1 = 1
    Call LabelMenu(1, rv1, rv0)
    ClosePipe
             
    rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
    'If rv2 = 1 Then
    '    rv2 = CBWTest_New_128_Sector_AU6377(2, 1)  ' write
    'End If
    
    If rv2 = 1 Then
        For i = 0 To 1
            rv2 = Write_Data_AU6377(LBA + i, 2, 65536)
            If rv2 <> 1 Then
                Exit For
            End If
        Next
    End If
    
    If rv2 = 1 Then
        For i = 0 To 1
            rv2 = Read_Data(LBA + i, 2, 65536)
            
            If rv2 <> 1 Then
                Exit For
            End If
            
            For j = 0 To 65535
                If Pattern_AU6377(j) <> ReadData(j) Then
                    Tester.Print "LBA= " & LBA + i & " Cycle= " & j & " Value= " & Hex(ReadData(j))
                    rv2 = 3
                    Exit For
                End If
            Next
        Next
    End If
    
    Call LabelMenu(2, rv2, rv1)
               
    ClosePipe
    If rv1 = 1 And rv2 <> 1 Then
        Tester.Label9.Caption = "SMC FAIL"
    End If
    
'============= SMC test begin =======================================
                 
    '--- for SMC
    CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
    Call MsecDelay(0.1)
                
    CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
    Call MsecDelay(0.1)
    ClosePipe
               
    rv3 = CBWTest_New(2, rv2, VidName)
    Call LabelMenu(2, rv3, rv2)
    ClosePipe
    
    If rv2 = 1 And rv3 <> 1 Then
        Tester.Label9.Caption = "XD FAIL"
    End If
                
'=============== SMC test END ==================================================
               
              
    rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
               
    If rv4 = 1 Then
        rv4 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
        If rv4 <> 1 Then
            Tester.Print "MS bus width Fail"
            rv4 = 2
        End If
    End If
            
    ClosePipe
    Call LabelMenu(3, rv4, rv3)
               
 ' HID mode and reader mode ---> compositive device
    
    If rv4 = 1 Then
        CardResult = DO_WritePort(card, Channel_P1A, &H7F)
        Call MsecDelay(0.2)
        If GetDeviceName("vid") <> "" Then
            rv4 = 0
            Tester.Print "NB mode test Fail"
            Call LabelMenu(3, rv4, rv3)
        Else
            Tester.Print "NBMD Test PASS"
        End If
        'CardResult = DO_WritePort(card, Channel_P1A, &H8)
        'Call MsecDelay(0.2)
    End If
            
 '======================== light test ======================
               
                   
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
    
    If rv4 = 1 Then
        'CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off  for HID mode
        'result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
        'Call MsecDelay(0.4)
          
        CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
        'Call MsecDelay(1.2)
        Call MsecDelay(0.1)
        rv4 = WaitDevOn(VidName)
        Call MsecDelay(0.2)
        CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
        
        If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
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
            GoTo AU6377ALFResult
        End If
                     
        '=======================
        '  key press test, it will return 8 when key up, GPI 6 must do low go hi action
        '========================
                
        Do
            CardResult = DO_WritePort(card, Channel_P1A, &H40) 'GPI6 : bit 6: pull high
            Sleep (200)
            CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' GPI6 : bit 6: pull low
            Sleep (400)
                       
            ReturnValue = fnInquiryBtnStatus(DeviceHandle)
            Tester.Print i; Space(5); "Key press value="; ReturnValue
            i = i + 1
        Loop While i < 3 And ReturnValue <> 10
            ' fnFreeDeviceHandle (DeviceHandle)
            'fnFreeDeviceHandle (DeviceHandle)
                     
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
                
                
AU6377ALFResult:
    CardResult = DO_WritePort(card, Channel_P1A, &H80)
    WaitDevOFF (VidName)
                        
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
    ElseIf rv4 = WRITE_FAIL Then
        MSWriteFail = MSWriteFail + 1
        TestResult = "MS_WF"
    ElseIf rv4 = READ_FAIL Then
        MSReadFail = MSReadFail + 1
        TestResult = "MS_RF"
    ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
                           
End Sub
Public Sub AU6476WLF35TestSub()
'AU6476BLF24 add MS bus width test, and modify SD bus width
'2011/7/4 AU6476BLF25 enhance XD 64k pattern R/W test
'2011/12/8 Add OpenShort test (GSMC)

On Error Resume Next

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
    GoTo AU6377ALFResult
End If

CardResult = DO_WritePort(card, Channel_P1C, &HFF)


Dim TmpChip As String
Dim RomSelector As Byte
   
TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
Tester.Print "Begin AU6476WL FT2 Test ..."
                    
'======================== Begin test ============================================
                  
    'Call MsecDelay(1.3)
                
    'If GetDeviceName("vid") <> "" Then
    '    rv0 = 0
    '    Tester.Print "NB mode test Fail"
    '    GoTo AU6377ALFResult
    'End If
               
    LBA = LBA + 1
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
                            
                
    '//////////////////////////////////////////////////
    '
    '   no card insert
    '
    '/////////////////////////////////////////////////
                
    Dim VidName As String
    Tester.Print LBA
    VidName = "vid_058f"
              
             
 '================================= Test light off =============================
                   
    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0000 1000
    Call MsecDelay(0.1)
    rv0 = WaitDevOn(VidName)
    Call MsecDelay(0.2)
                 
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '
    '  R/W test
    '
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                 
                
'initial return value
                
    If rv0 = 1 Then
        ClosePipe
        rv0 = CBWTest_New(0, 1, VidName)    ' sd slot
    End If
    
    If rv0 = 1 Then
        rv0 = Read_SD_Speed_AU6476_48MHz(0, 0, 64)
        ClosePipe
        
        If rv0 <> 1 Then
            Tester.Print "SD bus width Fail"
            rv0 = 2
        End If
    End If
    
    If rv0 = 1 Then
        rv0 = CBWTest_New_128_Sector_AU6377(0, rv0)  ' write
        ClosePipe
    End If
                
    Call LabelMenu(0, rv0, 1)
    
    
    
    rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
    ClosePipe
    Call LabelMenu(1, rv1, rv0)
   
   
             
    rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
    ClosePipe
    If rv2 = 1 Then
        rv2 = CBWTest_New_128_Sector_AU6377(2, 1)  ' write
        ClosePipe
    End If
    Call LabelMenu(2, rv2, rv1)
               
    
    If rv1 = 1 And rv2 <> 1 Then
        Tester.Label9.Caption = "SMC FAIL"
    End If
    
'============= SMC test begin =======================================
                 
    '--- for SMC
    CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
    Call MsecDelay(0.1)
                
    CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
    Call MsecDelay(0.1)
    ClosePipe
               
    rv3 = CBWTest_New(2, rv2, VidName)
    Call LabelMenu(2, rv3, rv2)
    ClosePipe
    
    If rv2 = 1 And rv3 <> 1 Then
        Tester.Label9.Caption = "XD FAIL"
    End If
                
'=============== SMC test END ==================================================
               
              
    rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
               
    If rv4 = 1 Then
        rv4 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
        If rv4 <> 1 Then
            Tester.Print "MS bus width Fail"
            rv4 = 2
        End If
    End If
            
    ClosePipe
    Call LabelMenu(3, rv4, rv3)
            
    If rv4 = 1 Then
        CardResult = DO_WritePort(card, Channel_P1A, &H7F)
        Call MsecDelay(0.2)
        If GetDeviceName("vid") <> "" Then
            rv4 = 0
            Tester.Print "NB mode test Fail"
            Call LabelMenu(3, rv4, rv3)
        Else
            Tester.Print "NBMD Test PASS"
        End If
        'CardResult = DO_WritePort(card, Channel_P1A, &H8)
        'Call MsecDelay(0.2)
    End If
            
 '======================== light test ======================
               
                   
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
    
    If rv4 = 1 Then
        'CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off  for HID mode
        'result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
        'Call MsecDelay(0.4)
          
        CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
        
        'Call MsecDelay(1.2)
        Call MsecDelay(0.1)
        rv4 = WaitDevOn(VidName)
        Call MsecDelay(0.2)
        CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
        
        If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
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
            GoTo AU6377ALFResult
        End If
                     
        '=======================
        '  key press test, it will return 8 when key up, GPI 6 must do low go hi action
        '========================
                
        Do
            CardResult = DO_WritePort(card, Channel_P1A, &H40) 'GPI6 : bit 6: pull high
            Sleep (200)
            CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' GPI6 : bit 6: pull low
            Sleep (1000)
                       
            ReturnValue = fnInquiryBtnStatus(DeviceHandle)
            Tester.Print i; Space(5); "Key press value="; ReturnValue
            i = i + 1
        Loop While i < 3 And ReturnValue <> 10
            ' fnFreeDeviceHandle (DeviceHandle)
            'fnFreeDeviceHandle (DeviceHandle)
                     
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
                
                
AU6377ALFResult:
     
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
    ElseIf rv4 = WRITE_FAIL Then
        MSWriteFail = MSWriteFail + 1
        TestResult = "MS_WF"
    ElseIf rv4 = READ_FAIL Then
        MSReadFail = MSReadFail + 1
        TestResult = "MS_RF"
    ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
                           
End Sub
Public Sub AU6476WLF36TestSub()
'AU6476BLF24 add MS bus width test, and modify SD bus width
'2011/7/4 AU6476BLF25 enhance XD 64k pattern R/W test
'2011/12/8 Add OpenShort test (GSMC)

On Error Resume Next

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
    GoTo AU6377ALFResult
End If

CardResult = DO_WritePort(card, Channel_P1C, &HFF)


Dim TmpChip As String
Dim RomSelector As Byte
Dim i As Long
Dim j As Long

TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
Tester.Print "Begin AU6476WL FT2 Test ..."
                    
'======================== Begin test ============================================
                  
    'Call MsecDelay(1.3)
                
    'If GetDeviceName("vid") <> "" Then
    '    rv0 = 0
    '    Tester.Print "NB mode test Fail"
    '    GoTo AU6377ALFResult
    'End If
               
    LBA = LBA + 1
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
                            
                
    '//////////////////////////////////////////////////
    '
    '   no card insert
    '
    '/////////////////////////////////////////////////
                
    Dim VidName As String
    Tester.Print LBA
    VidName = "vid_058f"
              
             
 '================================= Test light off =============================
                   
    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0000 1000
    'Call MsecDelay(0.1)
    rv0 = WaitDevOn(VidName)
    Call MsecDelay(0.2)
                 
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '
    '  R/W test
    '
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                 
                
'initial return value
                
    If rv0 = 1 Then
        ClosePipe
        rv0 = CBWTest_New(0, 1, VidName)    ' sd slot
    End If
    
    If rv0 = 1 Then
        rv0 = Read_SD_Speed_AU6476_48MHz(0, 0, 64)
        ClosePipe
        
        If rv0 <> 1 Then
            Tester.Print "SD bus width Fail"
            rv0 = 2
        End If
    End If
    
    If rv0 = 1 Then
        rv0 = CBWTest_New_128_Sector_AU6377(0, rv0)  ' write
        ClosePipe
    End If
                
    Call LabelMenu(0, rv0, 1)
    
    
    
    rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
    ClosePipe
    Call LabelMenu(1, rv1, rv0)
   
   
             
    rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
    
    If rv2 = 1 Then
        For i = 0 To 1
            rv2 = Write_Data_AU6377(LBA + i, 2, 65536)
            If rv2 <> 1 Then
                Exit For
            End If
        Next
    End If
    
    If rv2 = 1 Then
        For i = 0 To 1
            rv2 = Read_Data(LBA + i, 2, 65536)
            
            If rv2 <> 1 Then
                Exit For
            End If
            
            For j = 0 To 65535
                If Pattern_AU6377(j) <> ReadData(j) Then
                    Tester.Print "LBA= " & LBA + i & " Cycle= " & j & " Value= " & Hex(ReadData(j))
                    rv2 = 3
                    Exit For
                End If
            Next
        Next
    End If
    
    ClosePipe
    
    Call LabelMenu(2, rv2, rv1)
               
    
    If rv1 = 1 And rv2 <> 1 Then
        Tester.Label9.Caption = "SMC FAIL"
    End If
    
'============= SMC test begin =======================================
                 
    '--- for SMC
    CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
    Call MsecDelay(0.1)
                
    CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
    Call MsecDelay(0.1)
    ClosePipe
               
    rv3 = CBWTest_New(2, rv2, VidName)
    Call LabelMenu(2, rv3, rv2)
    ClosePipe
    
    If rv2 = 1 And rv3 <> 1 Then
        Tester.Label9.Caption = "XD FAIL"
    End If
                
'=============== SMC test END ==================================================
               
              
    rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
               
    If rv4 = 1 Then
        rv4 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
        If rv4 <> 1 Then
            Tester.Print "MS bus width Fail"
            rv4 = 2
        End If
    End If
            
    ClosePipe
    Call LabelMenu(3, rv4, rv3)
            
    If rv4 = 1 Then
        CardResult = DO_WritePort(card, Channel_P1A, &H7F)
        Call MsecDelay(0.2)
        If GetDeviceName("vid") <> "" Then
            rv4 = 0
            Tester.Print "NB mode test Fail"
            Call LabelMenu(3, rv4, rv3)
        Else
            Tester.Print "NBMD Test PASS"
        End If
        'CardResult = DO_WritePort(card, Channel_P1A, &H8)
        'Call MsecDelay(0.2)
    End If
            
 '======================== light test ======================
               
                   
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
    
    If rv4 = 1 Then
        'CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off  for HID mode
        'result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
        'Call MsecDelay(0.4)
          
        CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
        
        'Call MsecDelay(1.2)
        Call MsecDelay(0.1)
        rv4 = WaitDevOn(VidName)
        Call MsecDelay(0.2)
        CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
        
        If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
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
            GoTo AU6377ALFResult
        End If
                     
        '=======================
        '  key press test, it will return 8 when key up, GPI 6 must do low go hi action
        '========================
                
        Do
            CardResult = DO_WritePort(card, Channel_P1A, &H40) 'GPI6 : bit 6: pull high
            Sleep (200)
            CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' GPI6 : bit 6: pull low
            Sleep (400)
                       
            ReturnValue = fnInquiryBtnStatus(DeviceHandle)
            Tester.Print i; Space(5); "Key press value="; ReturnValue
            i = i + 1
        Loop While i < 3 And ReturnValue <> 10
            ' fnFreeDeviceHandle (DeviceHandle)
            'fnFreeDeviceHandle (DeviceHandle)
                     
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
                
                
AU6377ALFResult:
     
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
    WaitDevOFF (VidName)
    
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
    ElseIf rv4 = WRITE_FAIL Then
        MSWriteFail = MSWriteFail + 1
        TestResult = "MS_WF"
    ElseIf rv4 = READ_FAIL Then
        MSReadFail = MSReadFail + 1
        TestResult = "MS_RF"
    ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
                           
End Sub

Public Sub AU6476WLF3ETestSub()
'AU6476BLF24 add MS bus width test, and modify SD bus width
'2011/7/4 AU6476BLF25 enhance XD 64k pattern R/W test
'2011/12/8 Add OpenShort test (CSMC)

On Error Resume Next

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
    GoTo AU6377ALFResult
End If

CardResult = DO_WritePort(card, Channel_P1C, &HFF)


Dim TmpChip As String
Dim RomSelector As Byte

   
TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
Tester.Print "Begin AU6476WL FT2 Test ..."
                    
'======================== Begin test ============================================
                  
    'Call MsecDelay(1.3)
                
    'If GetDeviceName("vid") <> "" Then
    '    rv0 = 0
    '    Tester.Print "NB mode test Fail"
    '    GoTo AU6377ALFResult
    'End If
               
    LBA = LBA + 1
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
                            
                
    '//////////////////////////////////////////////////
    '
    '   no card insert
    '
    '/////////////////////////////////////////////////
                
    Dim VidName As String
    Tester.Print LBA
    VidName = "vid_058f"
              
             
 '================================= Test light off =============================
                   
    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0000 1000
    Call MsecDelay(0.1)
    rv0 = WaitDevOn(VidName)
    Call MsecDelay(0.2)
                 
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '
    '  R/W test
    '
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                 
                
'initial return value
                
    If rv0 = 1 Then
        ClosePipe
        rv0 = CBWTest_New(0, 1, VidName)    ' sd slot
    End If
    
    If rv0 = 1 Then
        rv0 = Read_SD_Speed_AU6476_48MHz(0, 0, 64)
        ClosePipe
        
        If rv0 <> 1 Then
            Tester.Print "SD bus width Fail"
            rv0 = 2
        End If
    End If
    
    If rv0 = 1 Then
        rv0 = CBWTest_New_128_Sector_AU6377(0, rv0)  ' write
        ClosePipe
    End If
                
    Call LabelMenu(0, rv0, 1)
    
    
    
    rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
    ClosePipe
    Call LabelMenu(1, rv1, rv0)
   
   
             
    rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
    ClosePipe
    If rv2 = 1 Then
        rv2 = CBWTest_New_128_Sector_AU6377(2, 1)  ' write
        ClosePipe
    End If
    Call LabelMenu(2, rv2, rv1)
               
    
    If rv1 = 1 And rv2 <> 1 Then
        Tester.Label9.Caption = "SMC FAIL"
    End If
    
'============= SMC test begin =======================================
                 
    '--- for SMC
    CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
    Call MsecDelay(0.1)
                
    CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
    Call MsecDelay(0.1)
    ClosePipe
               
    rv3 = CBWTest_New(2, rv2, VidName)
    Call LabelMenu(2, rv3, rv2)
    ClosePipe
    
    If rv2 = 1 And rv3 <> 1 Then
        Tester.Label9.Caption = "XD FAIL"
    End If
                
'=============== SMC test END ==================================================
               
              
    rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
               
    If rv4 = 1 Then
        rv4 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
        If rv4 <> 1 Then
            Tester.Print "MS bus width Fail"
            rv4 = 2
        End If
    End If
            
    ClosePipe
    Call LabelMenu(3, rv4, rv3)
            
    If rv4 = 1 Then
        CardResult = DO_WritePort(card, Channel_P1A, &H7F)
        Call MsecDelay(0.2)
        If GetDeviceName("vid") <> "" Then
            rv4 = 0
            Tester.Print "NB mode test Fail"
            Call LabelMenu(3, rv4, rv3)
        Else
            Tester.Print "NBMD Test PASS"
        End If
        'CardResult = DO_WritePort(card, Channel_P1A, &H8)
        'Call MsecDelay(0.2)
    End If
            
 '======================== light test ======================
               
                   
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
    
    If rv4 = 1 Then
        'CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off  for HID mode
        'result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
        'Call MsecDelay(0.4)
          
        CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
        
        'Call MsecDelay(1.2)
        Call MsecDelay(0.1)
        rv4 = WaitDevOn(VidName)
        Call MsecDelay(0.2)
        CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
        
        If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
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
            GoTo AU6377ALFResult
        End If
                     
        '=======================
        '  key press test, it will return 8 when key up, GPI 6 must do low go hi action
        '========================
                
        Do
            CardResult = DO_WritePort(card, Channel_P1A, &H40) 'GPI6 : bit 6: pull high
            Sleep (200)
            CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' GPI6 : bit 6: pull low
            Sleep (1000)
                       
            ReturnValue = fnInquiryBtnStatus(DeviceHandle)
            Tester.Print i; Space(5); "Key press value="; ReturnValue
            i = i + 1
        Loop While i < 3 And ReturnValue <> 10
            ' fnFreeDeviceHandle (DeviceHandle)
            'fnFreeDeviceHandle (DeviceHandle)
                     
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
                
                
AU6377ALFResult:
     
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
    ElseIf rv4 = WRITE_FAIL Then
        MSWriteFail = MSWriteFail + 1
        TestResult = "MS_WF"
    ElseIf rv4 = READ_FAIL Then
        MSReadFail = MSReadFail + 1
        TestResult = "MS_RF"
    ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
                           
End Sub
Public Sub AU6476WLF3FTestSub()
'AU6476BLF24 add MS bus width test, and modify SD bus width
'2011/7/4 AU6476BLF25 enhance XD 64k pattern R/W test
'2011/12/8 Add OpenShort test (CSMC)

On Error Resume Next

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
    GoTo AU6377ALFResult
End If

CardResult = DO_WritePort(card, Channel_P1C, &HFF)


Dim TmpChip As String
Dim RomSelector As Byte
Dim i As Long
Dim j As Long

TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
Tester.Print "Begin AU6476WL FT2 Test ..."
                    
'======================== Begin test ============================================
                  
    'Call MsecDelay(1.3)
                
    'If GetDeviceName("vid") <> "" Then
    '    rv0 = 0
    '    Tester.Print "NB mode test Fail"
    '    GoTo AU6377ALFResult
    'End If
               
    LBA = LBA + 1
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
                            
                
    '//////////////////////////////////////////////////
    '
    '   no card insert
    '
    '/////////////////////////////////////////////////
                
    Dim VidName As String
    Tester.Print LBA
    VidName = "vid_058f"
              
             
 '================================= Test light off =============================
                   
    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0000 1000
    'Call MsecDelay(0.1)
    rv0 = WaitDevOn(VidName)
    Call MsecDelay(0.2)
                 
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '
    '  R/W test
    '
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                 
                
'initial return value
                
    If rv0 = 1 Then
        ClosePipe
        rv0 = CBWTest_New(0, 1, VidName)    ' sd slot
    End If
    
    If rv0 = 1 Then
        rv0 = Read_SD_Speed_AU6476_48MHz(0, 0, 64)
        ClosePipe
        
        If rv0 <> 1 Then
            Tester.Print "SD bus width Fail"
            rv0 = 2
        End If
    End If
    
    If rv0 = 1 Then
        rv0 = CBWTest_New_128_Sector_AU6377(0, rv0)  ' write
        ClosePipe
    End If
                
    Call LabelMenu(0, rv0, 1)
    
    
    
    rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
    ClosePipe
    Call LabelMenu(1, rv1, rv0)
   
   
             
    rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
    
    If rv2 = 1 Then
        For i = 0 To 1
            rv2 = Write_Data_AU6377(LBA + i, 2, 65536)
            If rv2 <> 1 Then
                Exit For
            End If
        Next
    End If
    
    If rv2 = 1 Then
        For i = 0 To 1
            rv2 = Read_Data(LBA + i, 2, 65536)
            
            If rv2 <> 1 Then
                Exit For
            End If
            
            For j = 0 To 65535
                If Pattern_AU6377(j) <> ReadData(j) Then
                    Tester.Print "LBA= " & LBA + i & " Cycle= " & j & " Value= " & Hex(ReadData(j))
                    rv2 = 3
                    Exit For
                End If
            Next
        Next
    End If
    
    ClosePipe
    
    Call LabelMenu(2, rv2, rv1)
               
    
    If rv1 = 1 And rv2 <> 1 Then
        Tester.Label9.Caption = "SMC FAIL"
    End If
    
'============= SMC test begin =======================================
                 
    '--- for SMC
    CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
    Call MsecDelay(0.1)
                
    CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
    Call MsecDelay(0.1)
    ClosePipe
               
    rv3 = CBWTest_New(2, rv2, VidName)
    Call LabelMenu(2, rv3, rv2)
    ClosePipe
    
    If rv2 = 1 And rv3 <> 1 Then
        Tester.Label9.Caption = "XD FAIL"
    End If
                
'=============== SMC test END ==================================================
               
              
    rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
               
    If rv4 = 1 Then
        rv4 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
        If rv4 <> 1 Then
            Tester.Print "MS bus width Fail"
            rv4 = 2
        End If
    End If
            
    ClosePipe
    Call LabelMenu(3, rv4, rv3)
            
    If rv4 = 1 Then
        CardResult = DO_WritePort(card, Channel_P1A, &H7F)
        Call MsecDelay(0.2)
        If GetDeviceName("vid") <> "" Then
            rv4 = 0
            Tester.Print "NB mode test Fail"
            Call LabelMenu(3, rv4, rv3)
        Else
            Tester.Print "NBMD Test PASS"
        End If
        'CardResult = DO_WritePort(card, Channel_P1A, &H8)
        'Call MsecDelay(0.2)
    End If
            
 '======================== light test ======================
               
                   
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
    
    If rv4 = 1 Then
        'CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off  for HID mode
        'result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
        'Call MsecDelay(0.4)
          
        CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
        'Call MsecDelay(1.2)
        Call MsecDelay(0.1)
        rv4 = WaitDevOn(VidName)
        Call MsecDelay(0.2)
        CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
        
        If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
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
            GoTo AU6377ALFResult
        End If
                     
        '=======================
        '  key press test, it will return 8 when key up, GPI 6 must do low go hi action
        '========================
                
        Do
            CardResult = DO_WritePort(card, Channel_P1A, &H40) 'GPI6 : bit 6: pull high
            Sleep (200)
            CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' GPI6 : bit 6: pull low
            Sleep (1000)
                       
            ReturnValue = fnInquiryBtnStatus(DeviceHandle)
            Tester.Print i; Space(5); "Key press value="; ReturnValue
            i = i + 1
        Loop While i < 3 And ReturnValue <> 10
            ' fnFreeDeviceHandle (DeviceHandle)
            'fnFreeDeviceHandle (DeviceHandle)
                     
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
                
                
AU6377ALFResult:
     
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
    ElseIf rv4 = WRITE_FAIL Then
        MSWriteFail = MSWriteFail + 1
        TestResult = "MS_WF"
    ElseIf rv4 = READ_FAIL Then
        MSReadFail = MSReadFail + 1
        TestResult = "MS_RF"
    ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
                           
End Sub
Public Sub AU6476WLF05TestSub()
'AU6476BLF24 add MS bus width test, and modify SD bus width
'2011/7/4 AU6476BLF25 enhance XD 64k pattern R/W test

On Error Resume Next

If PCI7248InitFinish = 0 Then
    PCI7248ExistAU6254
    Call SetTimer_1ms
End If


Dim TmpChip As String
Dim RomSelector As Byte
Dim HV_Flag As Boolean      '1 = HV, 0 = LV
Dim HV_Res As String
Dim LV_Res As String
Dim Retry As Integer
    

TmpRes = ""
TmpChip = ChipName
'==================================== Switch assign ==========================================
            


HV_Flag = False

'======================== Begin test ============================================
                  
    'Call MsecDelay(1.3)
                
    'If GetDeviceName("vid") <> "" Then
    '    rv0 = 0
    '    Tester.Print "NB mode test Fail"
    '    GoTo AU6377ALFResult
    'End If

Routine_Label:


If Not HV_Flag Then
    Call PowerSet2(1, "5.3", "0.5", 1, "5.3", "0.5", 1)
    Tester.Print "Begin AU6476WL HV(5.3V) Test ..."
Else
    Call PowerSet2(1, "4.7", "0.5", 1, "4.7", "0.5", 1)
    Tester.Print "Begin AU6476WL LV(4.7V) Test ..."
End If

    TestResult = ""
    LBA = LBA + 1
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
                            
                
    '//////////////////////////////////////////////////
    '
    '   no card insert
    '
    '/////////////////////////////////////////////////
                
    Dim VidName As String
    Tester.Print LBA
    VidName = "vid_058f"
              
             
 '================================= Test light off =============================
                   
    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0000 1000
    Call MsecDelay(0.2)
    rv0 = WaitDevOn(VidName)
    Call MsecDelay(0.4)
                 
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '
    '  R/W test
    '
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                 
                
'initial return value
                
    If rv0 = 1 Then
        ClosePipe
        rv0 = CBWTest_New(0, 1, VidName)    ' sd slot
    End If
    
    If rv0 = 1 Then
        rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
        ClosePipe
        If rv0 <> 1 Then
            Tester.Print "SD bus width Fail"
            rv0 = 2
        End If
    End If
    
    If rv0 = 1 Then
        rv0 = CBWTest_New_128_Sector_AU6377(0, rv0)  ' write
        ClosePipe
    End If
                
    Call LabelMenu(0, rv0, 1)
    
    
    
    rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
    ClosePipe
    Call LabelMenu(1, rv1, rv0)
   
   
             
    rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
    ClosePipe
    If rv2 = 1 Then
        rv2 = CBWTest_New_128_Sector_AU6377(2, 1)  ' write
        ClosePipe
    End If
    Call LabelMenu(2, rv2, rv1)
               
    
    If rv1 = 1 And rv2 <> 1 Then
        Tester.Label9.Caption = "SMC FAIL"
    End If
    
'============= SMC test begin =======================================
                 
    '--- for SMC
    CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
    Call MsecDelay(0.1)
                
    CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
    Call MsecDelay(0.1)
    ClosePipe
               
    rv3 = CBWTest_New(2, rv2, VidName)
    Call LabelMenu(2, rv3, rv2)
    ClosePipe
    
    If rv2 = 1 And rv3 <> 1 Then
        Tester.Label9.Caption = "XD FAIL"
    End If
                
'=============== SMC test END ==================================================
               
              
    rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
               
    If rv4 = 1 Then
        rv4 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
        If rv4 <> 1 Then
            Tester.Print "MS bus width Fail"
            rv4 = 2
        End If
    End If
            
    ClosePipe
    Call LabelMenu(3, rv4, rv3)
            
    If rv4 = 1 Then
        CardResult = DO_WritePort(card, Channel_P1A, &H7F)
        Call MsecDelay(0.2)
        If GetDeviceName("vid") <> "" Then
            rv4 = 0
            Tester.Print "NB mode test Fail"
            Call LabelMenu(3, rv4, rv3)
        Else
            Tester.Print "NBMD Test PASS"
        End If
        'CardResult = DO_WritePort(card, Channel_P1A, &H8)
        'Call MsecDelay(0.2)
    End If
            
 '======================== light test ======================
               
                   
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
    
    If rv4 = 1 Then
        'CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off  for HID mode
        'result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
        'Call MsecDelay(0.4)
          
        CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
        
        'Call MsecDelay(1.2)
        Call MsecDelay(0.1)
        rv4 = WaitDevOn(VidName)
        Call MsecDelay(0.2)
        CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
        
        If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
            'UsbSpeedTestResult = GPO_FAIL
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
            GoTo AU6377ALFResult
        End If
                     
        '=======================
        '  key press test, it will return 8 when key up, GPI 6 must do low go hi action
        '========================
                
        Do
            CardResult = DO_WritePort(card, Channel_P1A, &H40) 'GPI6 : bit 6: pull high
            Sleep (200)
            CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' GPI6 : bit 6: pull low
            Sleep (1000)
                       
            ReturnValue = fnInquiryBtnStatus(DeviceHandle)
            Tester.Print i; Space(5); "Key press value="; ReturnValue
            i = i + 1
        Loop While i < 3 And ReturnValue <> 10
            ' fnFreeDeviceHandle (DeviceHandle)
            'fnFreeDeviceHandle (DeviceHandle)
                     
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
                
                
AU6377ALFResult:
     
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
     
     
    If Not HV_Flag Then
        If rv0 * rv1 * rv2 * rv3 * rv4 = 0 Then
            HV_Result = "Bin2"
            Tester.Print "HV Unknow"
        ElseIf rv0 * rv1 * rv2 * rv3 * rv4 <> 1 Then
            HV_Result = "Fail"
            Tester.Print "HV Fail"
        ElseIf rv0 * rv1 * rv2 * rv3 * rv4 = 1 Then
            HV_Result = "PASS"
            Tester.Print "HV PASS"
        End If
        
        HV_Flag = True
        Do
            Retry = Retry + 1
            Call MsecDelay(0.1)
        Loop Until ((Retry > 20) Or (GetDeviceName_NoReply(VidName) = ""))
        GoTo Routine_Label
    Else
        If rv0 * rv1 * rv2 * rv3 * rv4 = 0 Then
            LV_Result = "Bin2"
            Tester.Print "LV Unknow"
        ElseIf rv0 * rv1 * rv2 * rv3 * rv4 <> 1 Then
            LV_Result = "Fail"
            Tester.Print "LV Fail"
        ElseIf rv0 * rv1 * rv2 * rv3 * rv4 = 1 Then
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

Public Sub AU6476WLF06TestSub()
'AU6476BLF24 add MS bus width test, and modify SD bus width
'2011/7/4 AU6476BLF25 enhance XD 64k pattern R/W test

'On Error Resume Next

If PCI7248InitFinish = 0 Then
    PCI7248ExistAU6254
    Call SetTimer_1ms
End If


Dim TmpChip As String
Dim RomSelector As Byte
Dim HV_Flag As Boolean      '1 = HV, 0 = LV
Dim HV_Res As String
Dim LV_Res As String
Dim Retry As Integer
Dim i As Long
Dim j As Long

TmpRes = ""
TmpChip = ChipName
'==================================== Switch assign ==========================================
            


HV_Flag = False

'======================== Begin test ============================================
                  
    'Call MsecDelay(1.3)
                
    'If GetDeviceName("vid") <> "" Then
    '    rv0 = 0
    '    Tester.Print "NB mode test Fail"
    '    GoTo AU6377ALFResult
    'End If

Routine_Label:


If Not HV_Flag Then
    Call PowerSet2(1, "5.3", "0.5", 1, "5.3", "0.5", 1)
    Tester.Print "Begin AU6476WL HV(5.3V) Test ..."
Else
    Call PowerSet2(1, "4.7", "0.5", 1, "4.7", "0.5", 1)
    Tester.Print "Begin AU6476WL LV(4.7V) Test ..."
End If

    TestResult = ""
    LBA = LBA + 1
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
                            
                
    '//////////////////////////////////////////////////
    '
    '   no card insert
    '
    '/////////////////////////////////////////////////
                
    Dim VidName As String
    Tester.Print LBA
    VidName = "vid_058f"
              
             
 '================================= Test light off =============================
                   
    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0000 1000
    'Call MsecDelay(0.2)
    rv0 = WaitDevOn(VidName)
    Call MsecDelay(0.2)
                 
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '
    '  R/W test
    '
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                 
                
'initial return value
                
    If rv0 = 1 Then
        ClosePipe
        rv0 = CBWTest_New(0, 1, VidName)    ' sd slot
    End If
    
    If rv0 = 1 Then
        rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
        ClosePipe
        If rv0 <> 1 Then
            Tester.Print "SD bus width Fail"
            rv0 = 2
        End If
    End If
    
    If rv0 = 1 Then
        rv0 = CBWTest_New_128_Sector_AU6377(0, rv0)  ' write
        ClosePipe
    End If
                
    Call LabelMenu(0, rv0, 1)
    
    
    
    rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
    ClosePipe
    Call LabelMenu(1, rv1, rv0)
   
   
             
    rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
    
    If rv2 = 1 Then
        For i = 0 To 1
            rv2 = Write_Data_AU6377(LBA + i, 2, 65536)
            If rv2 <> 1 Then
                Exit For
            End If
        Next
    End If
    
    If rv2 = 1 Then
        For i = 0 To 1
            rv2 = Read_Data(LBA + i, 2, 65536)
            
            If rv2 <> 1 Then
                Exit For
            End If
            
            For j = 0 To 65535
                If (Pattern_AU6377(j) <> ReadData(j)) Then
                    Tester.Print "LBA= " & LBA + i & " Cycle= " & j & " Value= " & Hex(ReadData(j))
                    rv2 = 3
                    Exit For
                End If
            Next
        Next
    End If
    
    ClosePipe
    
    Call LabelMenu(2, rv2, rv1)
               
    
    If rv1 = 1 And rv2 <> 1 Then
        Tester.Label9.Caption = "SMC FAIL"
    End If
    
'============= SMC test begin =======================================
                 
    '--- for SMC
    CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
    Call MsecDelay(0.1)
                
    CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
    Call MsecDelay(0.1)
    ClosePipe
               
    rv3 = CBWTest_New(2, rv2, VidName)
    Call LabelMenu(2, rv3, rv2)
    ClosePipe
    
    If rv2 = 1 And rv3 <> 1 Then
        Tester.Label9.Caption = "XD FAIL"
    End If
                
'=============== SMC test END ==================================================
               
              
    rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
               
    If rv4 = 1 Then
        rv4 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
        If rv4 <> 1 Then
            Tester.Print "MS bus width Fail"
            rv4 = 2
        End If
    End If
            
    ClosePipe
    Call LabelMenu(3, rv4, rv3)
            
    If rv4 = 1 Then
        CardResult = DO_WritePort(card, Channel_P1A, &H7F)
        Call MsecDelay(0.2)
        If GetDeviceName("vid") <> "" Then
            rv4 = 0
            Tester.Print "NB mode test Fail"
            Call LabelMenu(3, rv4, rv3)
        Else
            Tester.Print "NBMD Test PASS"
        End If
    End If
            
 '======================== light test ======================
               
                   
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
    
    If rv4 = 1 Then
        'CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off  for HID mode
        'result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
        'Call MsecDelay(0.4)
          
        CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
        
        'Call MsecDelay(1.2)
        Call MsecDelay(0.1)
        rv4 = WaitDevOn(VidName)
        Call MsecDelay(0.2)
        CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
        
        If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
            'UsbSpeedTestResult = GPO_FAIL
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
            GoTo AU6377ALFResult
        End If
                     
        '=======================
        '  key press test, it will return 8 when key up, GPI 6 must do low go hi action
        '========================
                
        Do
            CardResult = DO_WritePort(card, Channel_P1A, &H40) 'GPI6 : bit 6: pull high
            Sleep (200)
            CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' GPI6 : bit 6: pull low
            Sleep (400)
                       
            ReturnValue = fnInquiryBtnStatus(DeviceHandle)
            Tester.Print i; Space(5); "Key press value="; ReturnValue
            i = i + 1
        Loop While i < 3 And ReturnValue <> 10
            ' fnFreeDeviceHandle (DeviceHandle)
            'fnFreeDeviceHandle (DeviceHandle)
                     
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
                
                
AU6377ALFResult:
     
    CardResult = DO_WritePort(card, Channel_P1A, &HFF)
    WaitDevOFF (VidName)
     
    If Not HV_Flag Then
        If rv0 * rv1 * rv2 * rv3 * rv4 = 0 Then
            HV_Result = "Bin2"
            Tester.Print "HV Unknow"
        ElseIf rv0 * rv1 * rv2 * rv3 * rv4 <> 1 Then
            HV_Result = "Fail"
            Tester.Print "HV Fail"
        ElseIf rv0 * rv1 * rv2 * rv3 * rv4 = 1 Then
            HV_Result = "PASS"
            Tester.Print "HV PASS"
        End If
        
        HV_Flag = True
        Do
            Retry = Retry + 1
            Call MsecDelay(0.1)
        Loop Until ((Retry > 20) Or (GetDeviceName_NoReply(VidName) = ""))
        GoTo Routine_Label
    Else
        If rv0 * rv1 * rv2 * rv3 * rv4 = 0 Then
            LV_Result = "Bin2"
            Tester.Print "LV Unknow"
        ElseIf rv0 * rv1 * rv2 * rv3 * rv4 <> 1 Then
            LV_Result = "Fail"
            Tester.Print "LV Fail"
        ElseIf rv0 * rv1 * rv2 * rv3 * rv4 = 1 Then
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
 Public Sub AU6476RLF25TestSub()
'AU6476BLF24 add MS bus width test, and modify SD bus width
' 981006 : add Open-short Test
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
Dim GPIPin As Byte
Dim LL As Long
               
   
TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
 
                
Tester.Print "AU6476RL is NB mode"
       
'CardResult = DO_WritePort(card, Channel_P1A, &H78)
 CardResult = DO_WritePort(card, Channel_P1A, &HC)
'CardResult = DO_WritePort(card, Channel_P1A, &HE) '6362 SD write protect is high active
    
'======================== Begin test ============================================
                  
                Call MsecDelay(1.2)
               
          
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
                Tester.Print LBA
                
                VidName = "058f"
              
                
       
                
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
                 rv0 = CBWTest_New(0, 1, VidName)    ' sd slot
                 
                
                 If rv0 = 1 Then
                    rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                    If rv0 <> 1 Then
                        Tester.Print "SD bus width Fail"
                    rv0 = 2
                    End If
                 End If
                 
                 
                 If rv0 = 1 Then
                  Tester.Print "===Read EEPROM ===="
                 rv0 = Read_EEPRomData(0, 8, 8)
                   If rv0 <> 1 Then
                    Tester.Print "EEPROMBus"
                    rv0 = 2
                    End If
                 
                 
                 End If
                ' test write protect
       
                    CardResult = DO_WritePort(card, Channel_P1A, &HF)
                   Call MsecDelay(0.1)
                    CardResult = DO_WritePort(card, Channel_P1A, &HE)
                   Call MsecDelay(0.1)
              
 
 
               rv0 = TestUnitReady(0)
 
              rv0 = RequestSense(0)

              rv0 = Write_Data_WPTest(0, 0, 2048)
                 
                 
                  If rv0 <> 1 Then
                  rv0 = 2
                  GoTo AU6377ALFResult
                 
                  End If
                 
                 
                    GPIPin = Read_GPI_AU6476(0, 0, 4)
                    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                     ClosePipe
                    Tester.Print "GPIPin:"; Hex(GPIPin)
                     Call MsecDelay(0.3)
                    Tester.Print "Port:"; Hex(LightOff)
                    
                        If GPIPin <> &H86 Or LightOff <> &HFE Then
                        Tester.Print "OS Fail"
                        rv0 = 3
                         Tester.Label9.Caption = "OS Test Fail"
                          GoTo AU6377ALFResult
                      End If
 
                Call LabelMenu(0, rv0, 1)
   
                
               '=============== SMC test END ==================================================
            ' CardResult = DO_WritePort(card, Channel_P1A, &H53)
            ' Call MsecDelay(1.2)
            
              rv4 = CBWTest_New(1, rv0, VidName)   'MS card test
               
               If rv4 = 1 Then
               rv4 = Read_MS_Speed_AU6476(1, 0, 64, "4Bits")
                 If rv4 <> 1 Then
                   Tester.Print "MS bus width Fail"
                   rv4 = 2
                 End If
              End If
            
              ClosePipe
               Call LabelMenu(3, rv4, rv0)
               
 '======================== light test ======================
               
               
               
               CardResult = DO_WritePort(card, Channel_P1A, &H3B)
  
               Call MsecDelay(0.2)
                
                If GetDeviceName("vid") <> "" Then
                     rv0 = 0
                    Tester.Print "NB mode test Fail"
                     GoTo AU6377ALFResult
                  
                End If
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H80)
  
                  Call MsecDelay(0.1)
                  
                  CardResult = DO_WritePort(card, Channel_P1A, &H77)
                 Call MsecDelay(1.4)
                If GetDeviceName("vid") = "" Then
                     rv0 = 0
                    Tester.Print "Normal mode test Fail"
                     GoTo AU6377ALFResult
                  
                End If
                
                
                
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
          
                Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print "LBA="; LBA
         
                
AU6377ALFResult:
                        CardResult = DO_WritePort(card, Channel_P1A, &H18)
  
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
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        
                        End If
                
               
                  
              
                
                
                
     '    CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off
     '    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
     '    Call MsecDelay(0.8)
     '    CardResult = DO_WritePort(card, Channel_P1A, &H40)
         
                
                
               
   End Sub
   Public Sub AU6476RLF26TestSub()
'AU6476BLF24 add MS bus width test, and modify SD bus width
' 981006 : add Open-short Test
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
Dim GPIPin As Byte
Dim LL As Long
               
   
TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
 
                
Tester.Print "AU6476RL is NB mode"
       
'CardResult = DO_WritePort(card, Channel_P1A, &H78)
 CardResult = DO_WritePort(card, Channel_P1A, &HC)
'CardResult = DO_WritePort(card, Channel_P1A, &HE) '6362 SD write protect is high active
    
'======================== Begin test ============================================
                  
                Call MsecDelay(1.2)
               
          
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
                Tester.Print LBA
                
                VidName = "058f"
              
                
       
                
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
                 rv0 = CBWTest_New(0, 1, VidName)    ' sd slot
                 
                
                 If rv0 = 1 Then
                    rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                    If rv0 <> 1 Then
                        Tester.Print "SD bus width Fail"
                    rv0 = 2
                    End If
                 End If
                 
                 
                 If rv0 = 1 Then
                  Tester.Print "===Read EEPROM ===="
                 rv0 = Read_EEPRomData(0, 8, 8)
                   If rv0 <> 1 Then
                    Tester.Print "EEPROMBus"
                    rv0 = 2
                    End If
                 
                 
                 End If
                ' test write protect
       
                    CardResult = DO_WritePort(card, Channel_P1A, &HF)
                   Call MsecDelay(0.1)
                    CardResult = DO_WritePort(card, Channel_P1A, &HE)
                   Call MsecDelay(0.1)
              
 
 
               rv0 = TestUnitReady(0)
 
              rv0 = RequestSense(0)

              rv0 = Write_Data_WPTest(0, 0, 2048)
                 
                 
                  If rv0 <> 1 Then
                  rv0 = 2
                  GoTo AU6377ALFResult
                 
                  End If
                 
                 
                    GPIPin = Read_GPI_AU6476(0, 0, 4)
                    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                     ClosePipe
                    Tester.Print "GPIPin:"; Hex(GPIPin)
                     Call MsecDelay(0.3)
                    Tester.Print "Port:"; Hex(LightOff)
                    
                        If GPIPin <> &H86 Or LightOff <> &HFE Then
                        Tester.Print "OS Fail"
                        rv0 = 3
                         Tester.Label9.Caption = "OS Test Fail"
                          GoTo AU6377ALFResult
                      End If
 
                Call LabelMenu(0, rv0, 1)
   
                
               '=============== SMC test END ==================================================
            ' CardResult = DO_WritePort(card, Channel_P1A, &H53)
            ' Call MsecDelay(1.2)
            
              rv4 = CBWTest_New(1, rv0, VidName)   'MS card test
               
               If rv4 = 1 Then
               rv4 = Read_MS_Speed_AU6476(1, 0, 64, "4Bits")
                 If rv4 <> 1 Then
                   Tester.Print "MS bus width Fail"
                   rv4 = 2
                 End If
              End If
            
              ClosePipe
               Call LabelMenu(3, rv4, rv0)
               
 '======================== light test ======================
               
               
               
               CardResult = DO_WritePort(card, Channel_P1A, &H3B)
  
               Call MsecDelay(0.2)
                
                If GetDeviceName("vid") <> "" Then
                     rv0 = 0
                    Tester.Print "NB mode test Fail"
                     GoTo AU6377ALFResult
                  
                End If
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H80)
  
                  Call MsecDelay(0.1)
                  
                  CardResult = DO_WritePort(card, Channel_P1A, &H77)
                 Call MsecDelay(1.4)
                If GetDeviceName("vid") = "" Then
                     rv0 = 0
                    Tester.Print "Normal mode test Fail"
                     GoTo AU6377ALFResult
                  
                End If
                
                
                
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
          
                Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print "LBA="; LBA
         
                
AU6377ALFResult:
                       CardResult = DO_WritePort(card, Channel_P1A, &HC)
  
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
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        
                        End If
                
               
                  
              
                
                
                
     '    CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off
     '    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
     '    Call MsecDelay(0.8)
     '    CardResult = DO_WritePort(card, Channel_P1A, &H40)
         
                
                
               
   End Sub
Public Sub AU6476RLF27TestSub()
'AU6476BLF24 add MS bus width test, and modify SD bus width
' 981006 : add Open-short Test
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
Dim GPIPin As Byte
Dim LL As Long
               
   
TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
 
                
Tester.Print "AU6476RL is NB mode"
       
'CardResult = DO_WritePort(card, Channel_P1A, &H78)
 CardResult = DO_WritePort(card, Channel_P1A, &H4C)
'CardResult = DO_WritePort(card, Channel_P1A, &HE) '6362 SD write protect is high active
    
'======================== Begin test ============================================
                  
                Call MsecDelay(1.2)
                
 CardResult = DO_WritePort(card, Channel_P1A, &HC)
 Call MsecDelay(0.2)
          

               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
                Tester.Print LBA
                
                VidName = "058f"
              
                
       
                
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
                 
                 rv0 = CBWTest_New(0, 1, VidName)    ' sd slot
                 
                
                 If rv0 = 1 Then
                    rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                    If rv0 <> 1 Then
                        Tester.Print "SD bus width Fail"
                    rv0 = 2
                    End If
                 End If
                 
                 
                 If rv0 = 1 Then
                  Tester.Print "===Read EEPROM ===="
                 rv0 = Read_EEPRomData(0, 8, 8)
                   If rv0 <> 1 Then
                    Tester.Print "EEPROMBus"
                    rv0 = 2
                    End If
                 
                 
                 End If
                ' test write protect
       
                    CardResult = DO_WritePort(card, Channel_P1A, &H4F)
                   Call MsecDelay(0.1)
                    CardResult = DO_WritePort(card, Channel_P1A, &H4E)
                   Call MsecDelay(0.1)
              
 
 
               rv0 = TestUnitReady(0)
 
              rv0 = RequestSense(0)

              rv0 = Write_Data_WPTest(0, 0, 2048)
                 
                 
                  If rv0 <> 1 Then
                  rv0 = 2
                  GoTo AU6377ALFResult
                 
                  End If
                 
                 
                    GPIPin = Read_GPI_AU6476(0, 0, 4)
                    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                     ClosePipe
                    Tester.Print "GPIPin:"; Hex(GPIPin)
                     Call MsecDelay(0.3)
                    Tester.Print "Port:"; Hex(LightOff)
                    
                        If (GPIPin <> &H86) Or (LightOff <> &HFD) Then
                            Tester.Print "OS Fail"
                            rv0 = 3
                            Tester.Label9.Caption = "OS Test Fail"
                            GoTo AU6377ALFResult
                        End If
 
                Call LabelMenu(0, rv0, 1)
   
                
               '=============== SMC test END ==================================================
            ' CardResult = DO_WritePort(card, Channel_P1A, &H53)
            ' Call MsecDelay(1.2)
            
              rv4 = CBWTest_New(1, rv0, VidName)   'MS card test
               
               If rv4 = 1 Then
               rv4 = Read_MS_Speed_AU6476(1, 0, 64, "4Bits")
                 If rv4 <> 1 Then
                   Tester.Print "MS bus width Fail"
                   rv4 = 2
                 End If
              End If
            
              ClosePipe
               Call LabelMenu(3, rv4, rv0)
               
 '======================== light test ======================
               
               
               
               CardResult = DO_WritePort(card, Channel_P1A, &H3B)
  
               Call MsecDelay(0.2)
                
                If GetDeviceName("vid") <> "" Then
                     rv0 = 0
                    Tester.Print "NB mode test Fail"
                     GoTo AU6377ALFResult
                  
                End If
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H80)
  
                  Call MsecDelay(0.1)
                  
                  CardResult = DO_WritePort(card, Channel_P1A, &H77)
                 Call MsecDelay(1.4)
                If GetDeviceName("vid") = "" Then
                     rv0 = 0
                    Tester.Print "Normal mode test Fail"
                     GoTo AU6377ALFResult
                  
                End If
                
                
                
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
          
                Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print "LBA="; LBA
         
                
AU6377ALFResult:
                       CardResult = DO_WritePort(card, Channel_P1A, &HC)
  
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
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        
                        End If
                
               
                  
              
                
                
                
     '    CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off
     '    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
     '    Call MsecDelay(0.8)
     '    CardResult = DO_WritePort(card, Channel_P1A, &H40)
         
                
                
               
   End Sub
Public Sub AU6476QLF25TestSub()
'AU6476BLF24 add MS bus width test, and modify SD bus width
' 981006 : add Open-short Test
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
Dim GPIPin As Byte
               
   
TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
Tester.Print "AU6476QL is NB mode"
       
CardResult = DO_WritePort(card, Channel_P1A, &H18)
  
    
'======================== Begin test ============================================
                  
                Call MsecDelay(1.2)
               
          
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
                Tester.Print LBA
                
                VidName = "058f"
              
                
       
                
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
                 rv0 = CBWTest_New(0, 1, VidName)    ' sd slot
                 
                
                 If rv0 = 1 Then
                    rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                    If rv0 <> 1 Then
                    Tester.Print "SD bus width Fail"
                    rv0 = 2
                    End If
                 End If
                 
                 
                 If rv0 = 1 Then
                  Tester.Print "===Read EEPROM ===="
                 rv0 = Read_EEPRomData(0, 8, 8)
                   If rv0 <> 1 Then
                    Tester.Print "EEPROMBus"
                    rv0 = 2
                    End If
                 
                 
                 End If
               ' test write protect
       
                  CardResult = DO_WritePort(card, Channel_P1A, &H19)
                  Call MsecDelay(0.1)
                  CardResult = DO_WritePort(card, Channel_P1A, &H1A)
                  Call MsecDelay(0.1)
              
 
 
              rv0 = TestUnitReady(0)
 
             rv0 = RequestSense(0)

             rv0 = Write_Data_WPTest(0, 0, 2048)
                 
                 
                 If rv0 <> 1 Then
                 rv0 = 2
                 GoTo AU6377ALFResult
                 
                 End If
                 
                 
                    GPIPin = Read_GPI_AU6476(0, 0, 4)
                    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                     ClosePipe
                    Tester.Print "GPIPin:"; Hex(GPIPin)
  
                    Tester.Print "Port:"; Hex(LightOff)
                    
                        If GPIPin <> &H7 Or LightOff <> &HFE Then
                        Tester.Print "OS Fail"
                        rv0 = 3
                         Tester.Label9.Caption = "OS Test Fail"
                          GoTo AU6377ALFResult
                      End If
 
                Call LabelMenu(0, rv0, 1)
   
                
               '=============== SMC test END ==================================================
 
              
               rv4 = CBWTest_New(3, rv0, VidName)   'MS card test
               
               If rv4 = 1 Then
               rv4 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                 If rv4 <> 1 Then
                   Tester.Print "MS bus width Fail"
                   rv4 = 2
                 End If
              End If
            
              ClosePipe
               Call LabelMenu(3, rv4, rv0)
               
 '======================== light test ======================
               
               
               
               CardResult = DO_WritePort(card, Channel_P1A, &H3F)
  
               Call MsecDelay(0.2)
                
                If GetDeviceName("vid") <> "" Then
                     rv0 = 0
                    Tester.Print "NB mode test Fail"
                     GoTo AU6377ALFResult
                  
                End If
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H80)
  
                  Call MsecDelay(0.1)
                  
                  CardResult = DO_WritePort(card, Channel_P1A, &H77)
                 Call MsecDelay(1.4)
                If GetDeviceName("vid") = "" Then
                     rv0 = 0
                    Tester.Print "Normal mode test Fail"
                     GoTo AU6377ALFResult
                  
                End If
                
                
                
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
          
                Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print "LBA="; LBA
         
                
AU6377ALFResult:
                        CardResult = DO_WritePort(card, Channel_P1A, &H18)
  
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
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        
                        End If
                
               
                  
              
                
                
                
     '    CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off
     '    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
     '    Call MsecDelay(0.8)
     '    CardResult = DO_WritePort(card, Channel_P1A, &H40)
         
                
                
               
   End Sub
   
 Public Sub AU6476LLF25TestSub()
'AU6476BLF24 add MS bus width test, and modify SD bus width
' 981006 : add Open-short Test
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
Dim GPIPin As Byte
               
   
TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
Tester.Print "AU6476LL is NB mode"
                  
 If Left(TmpChip, 8) = "AU6378AL" Or Left(TmpChip, 8) = "AU6476LL" Then
 
         CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  // force into reader mode
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.2)
         CardResult = DO_WritePort(card, Channel_P1A, &H0)
         
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H3F)
          
          
  End If
  
              
  
  
  '========================== AU6377 new board switch assign ment  ============
If TmpChip = "AU6377ALF21" Then ' this for new board  and internalrom
         RomSelector = &H10  '-------- this is for MS in pin
  End If
  
  
  If TmpChip = "AU6377ALF24" Then ' this for new board  and internalrom
         RomSelector = &H10
  End If
  
  If TmpChip = "AU6377ALF25" Then ' this for new board  and internalrom
         RomSelector = &H0
  End If
         
         
  If Left(TmpChip, 10) = "AU6377ALF2" Then
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector)  ' 1111 1110
         CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' 1111 1110
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector) ' 5th bit is rom selector, High is internal rom
  End If
  
  


  
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1.3)
               
                
                  If GetDeviceName("vid") <> "" Then
                     rv0 = 0
                    Tester.Print "NB mode test Fail"
                     GoTo AU6377ALFResult
                  
                   End If
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
                Tester.Print LBA
                
                   VidName = "058f"
               
  
              
             
 '================================= Test light off =============================
                
              
                
                
                   
                    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                  
                    
                   Call MsecDelay(1.3)
                 
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
                 rv0 = CBWTest_New(0, 1, VidName)    ' sd slot
                 
                
                 If rv0 = 1 Then
                    rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                    If rv0 <> 1 Then
                    Tester.Print "SD bus width Fail"
                    rv0 = 2
                    End If
                 End If
                 
               ' test write protect
       
                  CardResult = DO_WritePort(card, Channel_P1A, &H9)
                  Call MsecDelay(0.1)
                  CardResult = DO_WritePort(card, Channel_P1A, &HA)
                  Call MsecDelay(0.1)
              
 
 
              rv0 = TestUnitReady(0)
 
             rv0 = RequestSense(0)

             rv0 = Write_Data_WPTest(0, 0, 2048)
                 
                 
                 If rv0 <> 1 Then
                 rv0 = 2
                 GoTo AU6377ALFResult
                 
                 End If
                 
                 
                    GPIPin = Read_GPI_AU6476(0, 0, 4)
                    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                    
                    Tester.Print "GPIPin:"; Hex(GPIPin)
                    
                    
                    
                    Tester.Print "Port:"; Hex(LightOff)
                    
                        If GPIPin <> &H27 Or LightOff <> &HBE Then
                        Tester.Print "OS Fail"
                        rv0 = 3
                         Tester.Label9.Caption = "OS Test Fail"
                          GoTo AU6377ALFResult
                      End If
                       
                 
                 If rv0 = 1 And TmpChip = "AU6375HLF22" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                End If
                
                   If Left(TmpChip, 10) = "AU6377ALF2" Then
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                      LBA = TmpLBA
                   End If
                Call LabelMenu(0, rv0, 1)
              ClosePipe
            '      rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
                  ' rv1 = 1
            '    Call LabelMenu(1, rv1, rv0)
            '    ClosePipe
              
              
             
             '    rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
             '   Call LabelMenu(2, rv2, rv1)
               
             '   ClosePipe
                
             '    If rv1 = 1 And rv2 <> 1 Then
             '    Tester.Label9.Caption = "SMC FAIL"
             '   End If
                '============= SMC test begin =======================================
               
              
                 
                          '--- for SMC
              '  CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
              '  Call MsecDelay(0.1)
                      
                
              '  CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
              '  Call MsecDelay(0.1)
              '  ClosePipe
                
                
               
               '   rv3 = CBWTest_New(2, rv2, VidName)
               
              '  Call LabelMenu(2, rv3, rv2)
              '  ClosePipe
              '    If rv2 = 1 And rv3 <> 1 Then
              '   Tester.Label9.Caption = "XD FAIL"
               ' End If
                
                
              
                
               '=============== SMC test END ==================================================
 
              
               rv4 = CBWTest_New(3, rv0, VidName)   'MS card test
               
               If rv4 = 1 Then
               rv4 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                 If rv4 <> 1 Then
                   Tester.Print "MS bus width Fail"
                   rv4 = 2
                 End If
              End If
            
              ClosePipe
               Call LabelMenu(3, rv4, rv0)
               
 '======================== light test ======================
               
               
               
               
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
      If rv4 = 1 Then
          CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off  for HID mode
          result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
          Call MsecDelay(0.4)
          
          CardResult = DO_WritePort(card, Channel_P1A, &H7C) ' HID mode
                            
    
          Call MsecDelay(1.2)
          CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
          If rv4 = 1 And (LightOff <> &HBE) Then
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
                     
          Call MsecDelay(0.3)
                     DeviceHandle = &HFFFF  'invalid handle initial value
                     
                     ReturnValue = fnGetDeviceHandle(DeviceHandle)
                     Tester.Print ReturnValue; Space(5); ' 1: pass the other refer btnstatus.h
                     Tester.Print "DeviceHandle="; DevicehHandle
                     
                     If ReturnValue <> 1 Then
                           rv0 = UNKNOW       '---> HID mode unknow device mode
                           Call LabelMenu(0, rv0, 1)
                           Tester.Label9.Caption = "HID mode unknow device"
                          fnFreeDeviceHandle (DeviceHandle)
                         GoTo AU6377ALFResult
                     End If
                     
                     '=======================
                     '  key press test, it will return 8 when key up, GPI 6 must do low go hi action
                     '========================
                     
                
                     Do
                          CardResult = DO_WritePort(card, Channel_P1A, &H40) 'GPI6 : bit 6: pull high
                          Sleep (200)
                         CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' GPI6 : bit 6: pull low
                         Sleep (1000)
                       
                         ReturnValue = fnInquiryBtnStatus(DeviceHandle)
                         Tester.Print i; Space(5); "Key press value="; ReturnValue
                         i = i + 1
                     Loop While i < 3 And ReturnValue <> 10
                    ' fnFreeDeviceHandle (DeviceHandle)
                   '  fnFreeDeviceHandle (DeviceHandle)
                     
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
         
                
AU6377ALFResult:
                        
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
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        
                        End If
                
               
                  
              
                
                
                
     '    CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off
     '    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
     '    Call MsecDelay(0.8)
     '    CardResult = DO_WritePort(card, Channel_P1A, &H40)
         
                
                
               
   End Sub
   
Public Sub AU1111AAA10TestSub()
'AU6476BLF24 add MS bus width test, and modify SD bus width
' 981006 : add Open-short Test
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
Dim GPIPin As Byte
               
   
 
                 
  
              
             
 '================================= Test light off =============================
                
              
                
           
                  
                    
                   Call MsecDelay(1.3)
                 
                 '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                 '
                 '  R/W test
                 '
                 '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                 
                
                'initial return value
                
                
          
               
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
                
                
                ClosePipe
                 rv0 = CBWTest_New(0, 1, "vid")    ' sd slot
                 Dim ib As Byte
                 Tester.Print "===write ===="
                 For ib = 0 To 6
W1:                 rv0 = Write_EEPromData(0, ib + 8, ib)
                    Call MsecDelay(0.1)
                   If rv0 = 0 Then
                    GoTo W1
                   End If
           
                 Tester.Print ib, rv0
                 Next ib
                 
                   rv0 = Write_EEPromData(0, ib + 8, 28)
                   Call MsecDelay(0.1)
                   Tester.Print ib, rv0
                 If rv0 = 1 Then
                  Tester.Print "===Read ===="
                 rv0 = Read_EEPRomData(0, 8, 8)
                 End If
                
               If rv0 = 1 Then
                       Tester.Label9.Caption = "Write eeprom pass"
               Else
                       Tester.Label9.Caption = "Write eeprom fail"
               End If
             
AU6377ALFResult:
                        
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
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        
                        End If
                
               
                  
              
                
                
                
     '    CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off
     '    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
     '    Call MsecDelay(0.8)
     '    CardResult = DO_WritePort(card, Channel_P1A, &H40)
         
                
                
               
   End Sub
   
Public Sub AU6476JLO20TestSub()
'AU6476BLF24 add MS bus width test, and modify SD bus width
 
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
Dim VidName As String
   
 
            
If PCI7248InitFinish = 0 Then
      
      PCI7248_AU6476_ALPSExist
End If
                
Tester.Print "AU6476JL is NB mode"


VidName = "vid_058f"
                
LBA = LBA + 1
                
 '================================= Test light off =============================
 CardResult = DO_WritePort(card, Channel_P1A, &H18)      ' &H8 overcurrent, &H18 no overcurrent
 CardResult = DO_WritePort(card, Channel_P1CL, &H0)  ' send CF power
   ' CardResult = DO_WritePort(card, Channel_P1A, &H0)
Call MsecDelay(1.3)
                 
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
                
                     
                
             
                 rv0 = CBWTest_New(0, 1, VidName)    ' sd slot
                 
                 
                 If rv0 = 1 Then
                    rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                    If rv0 <> 1 Then
                    Tester.Print "SD bus width Fail"
                    rv0 = 2
                    End If
                 End If
                 
                 
                 
                
                Call LabelMenu(0, rv0, 1)
                ClosePipe
              
                  rv1 = CBWTest_NewCFOverCurrentRWnoPower(1, rv0, VidName)    ' cf slot
                  
                  If rv1 = 5 Then
                     Tester.Print "OverCurrent wrong action"
                  End If
                  ' rv1 = 1
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
              
             
                 rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
               
                ClosePipe
                
                 If rv1 = 1 And rv2 <> 1 Then
                 Tester.Label9.Caption = "SMC FAIL"
                End If
                '============= SMC test begin =======================================
               
              
                 
                          '--- for SMC
                CardResult = DO_WritePort(card, Channel_P1A, &H1C)   ' 0110 0100
                Call MsecDelay(0.1)
                      
                
                CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100 SMC high
                Call MsecDelay(0.1)
                ClosePipe
                
                
               
                  rv3 = CBWTest_New(2, rv2, VidName)
               
                Call LabelMenu(2, rv3, rv2)
                ClosePipe
                  If rv2 = 1 And rv3 <> 1 Then
                 Tester.Label9.Caption = "XD FAIL"
                End If
                
                
              
                
               '=============== SMC test END ==================================================
               
              
               rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
               
               If rv4 = 1 Then
               rv4 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                 If rv4 <> 1 Then
                   Tester.Print "MS bus width Fail"
                   rv4 = 2
                 End If
              End If
            
               ClosePipe
               
               'Call MsecDelay(0.7)
               
                    CardResult = DO_ReadPort(card, Channel_P1CH, LightOff)
          If rv4 = 1 And (LightOff <> 3) Then
                  UsbSpeedTestResult = GPO_FAIL
                    rv4 = 2
                    Tester.Print "GPO FAIL =" & LightOff
                    Tester.Label9.Caption = "GPO FAIL " & LightOff
           End If
          
  '=================================================================================
 
 ' over current protect test
 
 '=================================================================================
 
 If rv4 = 1 Then
                CardResult = DO_WritePort(card, Channel_P1A, &H16)   ' 0110 0100 remove CF, and do overcurrent
                Call MsecDelay(0.2)
                CardResult = DO_WritePort(card, Channel_P1A, &H14)      ' CF card in, and over current loading
               
                Call MsecDelay(0.2)
                 CardResult = DO_WritePort(card, Channel_P1A, &H4)   ' add loading
                 Call MsecDelay(0.2)
               CardResult = DO_WritePort(card, Channel_P1CL, &H1) ' change to loading
               
                  Tester.Print "Over current protest test"
             
                 rv5 = CBWTest_NewOverCurrent(1, 1, "058f")
                 ' Test CF at 1st slot
                '  CardResult = DO_WritePort(card, Channel_P1CL, &H0) ' change to loading
                ' CardResult = DO_WritePort(card, Channel_P1A, &H14) ' release loading
                 ': 6 means over upper spec
                 ': 5 means over lower spec
                 
                     If rv5 <> 3 Then
                         Tester.Print "over current protect test Fail"
                     End If
                     If rv5 = 3 Then
                     rv5 = 1
                     End If
                 CardResult = DO_WritePort(card, Channel_P1A, &H3F)
          
                      Call MsecDelay(1.4)
                     If GetDeviceName("vid") <> "" Then
                     rv0 = 0
                     rv5 = 2
                     Tester.Print "NB mode test Fail"
                     GoTo AU6377ALFResult
                  
                   End If
               
 
 End If
 
          
          
          
 '========================  overcurrent measeure ======================
               
               
               
               
            If rv5 = 1 Then
                   ' change to loading desing
                CardResult = DO_WritePort(card, Channel_P1A, &H80)    ' 0110 0100 remove CF, and do overcurrent
                
                Call MsecDelay(0.6)
                 
         
                
                CardResult = DO_WritePort(card, Channel_P1A, &H14)      ' CF card in, and over current loading
               
                Call MsecDelay(1.4)
                CardResult = DO_WritePort(card, Channel_P1CL, &H1)
                
              
               Tester.Print "A/D tester"
               AU6476Upper = 4.43
               AU6476Lower = 3.17
               ReaderExist = 0
                 rv6 = CBWTest_NewOverCurrent4(1, 1, "058f")     ' Test CF at 1st slot
                     
                 ': 6 means over upper spec
                 ': 5 means over lower spec
                 Tester.Print rv6, "6> 4.43, 5<3.17"
                    
                 
                    
               
           End If
               

 
 
 
 ' HID mode and reader mode ---> compositive device
      If rv6 = 1 Then
          CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off  for HID mode
           
          Call MsecDelay(0.4)
          
          CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
                            
    
          Call MsecDelay(1.5)
          
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
                         GoTo AU6377ALFResult
                     End If
                     
                     '=======================
                     '  key press test, it will return 8 when key up, GPI 6 must do low go hi action
                     '========================
                     
                
                     Do
                          CardResult = DO_WritePort(card, Channel_P1A, &H40) 'GPI6 : bit 6: pull high
                          Sleep (200)
                         CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' GPI6 : bit 6: pull low
                         Sleep (1000)
                       
                         ReturnValue = fnInquiryBtnStatus(DeviceHandle)
                         Tester.Print i; Space(5); "Key press value="; ReturnValue
                         i = i + 1
                     Loop While i < 3 And ReturnValue <> 10
                    ' fnFreeDeviceHandle (DeviceHandle)
                   '  fnFreeDeviceHandle (DeviceHandle)
                     
                     If ReturnValue <> 10 Then
                     
                      rv1 = 2
                     
                       Call LabelMenu(1, rv1, rv0)
                       Label9.Caption = "KeyPress Fail"
                       
                     End If
                              
                     '   CardResult = DO_WritePort(card, Channel_P1A, &H3F)
          
               
                 
           End If
                
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv5, " \\CF overcurrent protect"
                Tester.Print rv6, " \\CF power , 6: over upper limit, 5: under lower limit"
                Tester.Print "LBA="; LBA
                
AU6377ALFResult:
                       
                        
                       If rv6 = 5 Or rv6 = 6 Or rv5 <> 1 Then
                        
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                            
                            If rv0 * rv1 * rv2 * rv3 * rv4 = 1 Then
                              XDWriteFail = XDWriteFail + 1
                              TestResult = "XD_WF"
                            End If
                            
                            
                        ElseIf rv0 = UNKNOW Then
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
                        ElseIf rv2 = WRITE_FAIL Or rv3 = WRITE_FAIL Or rv4 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                             CFWriteFail = CFWriteFail + 1
                            TestResult = "CF_WF"
                        ElseIf rv2 = READ_FAIL Or rv3 = READ_FAIL Or rv4 = READ_FAIL Or rv5 = READ_FAIL Then
                            CFWriteFail = CFWriteFail + 1
                            TestResult = "CF_WF"
        
                        ElseIf rv5 * rv6 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        
                        End If
                
               
                  
              
                
                
                
     '    CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off
     '    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
     '    Call MsecDelay(0.8)
     '    CardResult = DO_WritePort(card, Channel_P1A, &H40)
         
                
                
               
   End Sub
   
   Public Sub AU6476JLO20TestSubOld()
'AU6476BLF24 add MS bus width test, and modify SD bus width
 
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
Dim VidName As String
   
 
            
If PCI7248_AU6476_ALPSExist = 0 Then
      PCI7248Exist
End If
                
Tester.Print "AU6476JL is NB mode"


VidName = "vid_058f"
                
LBA = LBA + 1
                
 '================================= Test light off =============================
 CardResult = DO_WritePort(card, Channel_P1A, &H18)      ' &H8 overcurrent, &H18 no overcurrent
 CardResult = DO_WritePort(card, Channel_P1CL, &H0)  ' send CF power
   ' CardResult = DO_WritePort(card, Channel_P1A, &H0)
Call MsecDelay(1.3)
                 
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
                
                     
                
             
                 rv0 = CBWTest_New(0, 1, VidName)    ' sd slot
                 
                 
                 If rv0 = 1 Then
                    rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                    If rv0 <> 1 Then
                    Tester.Print "SD bus width Fail"
                    rv0 = 2
                    End If
                 End If
                 
                 
                 
                
                Call LabelMenu(0, rv0, 1)
                ClosePipe
              
                  rv1 = CBWTest_NewCFOverCurrentRWnoPower(1, rv0, VidName)    ' cf slot
                  
                  If rv1 = 5 Then
                     Tester.Print "OverCurrent wrong action"
                  End If
                  ' rv1 = 1
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
              
             
                 rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
               
                ClosePipe
                
                 If rv1 = 1 And rv2 <> 1 Then
                 Tester.Label9.Caption = "SMC FAIL"
                End If
                '============= SMC test begin =======================================
               
              
                 
                          '--- for SMC
                CardResult = DO_WritePort(card, Channel_P1A, &H1C)   ' 0110 0100
                Call MsecDelay(0.1)
                      
                
                CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100 SMC high
                Call MsecDelay(0.1)
                ClosePipe
                
                
               
                  rv3 = CBWTest_New(2, rv2, VidName)
               
                Call LabelMenu(2, rv3, rv2)
                ClosePipe
                  If rv2 = 1 And rv3 <> 1 Then
                 Tester.Label9.Caption = "XD FAIL"
                End If
                
                
              
                
               '=============== SMC test END ==================================================
               
              
               rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
               
               If rv4 = 1 Then
               rv4 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                 If rv4 <> 1 Then
                   Tester.Print "MS bus width Fail"
                   rv4 = 2
                 End If
              End If
            
               ClosePipe
               
               'Call MsecDelay(0.7)
               
                    CardResult = DO_ReadPort(card, Channel_P1CH, LightOff)
          If rv4 = 1 And (LightOff <> 3 And LightOff <> 11 And LightOff <> 15) Then
                  UsbSpeedTestResult = GPO_FAIL
                    rv4 = 2
                    Tester.Print "GPO FAIL =" & LightOff
                    Tester.Label9.Caption = "GPO FAIL " & LightOff
           End If
          
  '=================================================================================
 
 ' over current protect test
 
 '=================================================================================
 
 If rv4 = 1 Then
                CardResult = DO_WritePort(card, Channel_P1A, &H16)   ' 0110 0100 remove CF, and do overcurrent
                Call MsecDelay(0.2)
                CardResult = DO_WritePort(card, Channel_P1A, &H14)      ' CF card in, and over current loading
               
                Call MsecDelay(0.2)
                 CardResult = DO_WritePort(card, Channel_P1A, &H4)   ' add loading
                 Call MsecDelay(0.2)
                CardResult = DO_WritePort(card, Channel_P1CL, &H1) ' change to loading
               
                  Tester.Print "Over current protest test"
             
                 rv5 = CBWTest_NewOverCurrent(1, 1, "058f")
                 ' Test CF at 1st slot
                '  CardResult = DO_WritePort(card, Channel_P1CL, &H0) ' change to loading
                ' CardResult = DO_WritePort(card, Channel_P1A, &H14) ' release loading
                 ': 6 means over upper spec
                 ': 5 means over lower spec
                 
                     If rv5 <> 3 Then
                         Tester.Print "over current protect test Fail"
                     End If
                     If rv5 = 3 Then
                     rv5 = 1
                     End If
                 CardResult = DO_WritePort(card, Channel_P1A, &H3F)
          
                     Call MsecDelay(0.6)
                    If GetDeviceName("vid") <> "" Then
                     rv0 = 0
                     rv5 = 2
                     Tester.Print "NB mode test Fail"
                     GoTo AU6377ALFResult
                  
                    End If
               
 
 End If
 
          
          
          
 '========================  overcurrent measeure ======================
               
               
               
               
            If rv5 = 1 Then
                   ' change to loading desing
                CardResult = DO_WritePort(card, Channel_P1A, &H80)    ' 0110 0100 remove CF, and do overcurrent
                
                Call MsecDelay(0.6)
                 
         
                
                CardResult = DO_WritePort(card, Channel_P1A, &H14)      ' CF card in, and over current loading
               
                Call MsecDelay(1.4)
                CardResult = DO_WritePort(card, Channel_P1CL, &H1)
                
              
               Tester.Print "A/D tester"
               AU6476Upper = 4.43
               AU6476Lower = 3.17
               ReaderExist = 0
                 rv6 = CBWTest_NewOverCurrent4(1, 1, "058f")     ' Test CF at 1st slot
                     
                 ': 6 means over upper spec
                 ': 5 means over lower spec
                 Tester.Print rv6, "6> 4.43, 5<3.17"
                    
                 
                    
               
           End If
               

 
 
 
 ' HID mode and reader mode ---> compositive device
      If rv6 = 1 Then
          CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off  for HID mode
           
          Call MsecDelay(0.4)
          
          CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
                            
    
          Call MsecDelay(1.5)
          
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
                         GoTo AU6377ALFResult
                     End If
                     
                     '=======================
                     '  key press test, it will return 8 when key up, GPI 6 must do low go hi action
                     '========================
                     
                
                     Do
                          CardResult = DO_WritePort(card, Channel_P1A, &H40) 'GPI6 : bit 6: pull high
                          Sleep (200)
                         CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' GPI6 : bit 6: pull low
                         Sleep (1000)
                       
                         ReturnValue = fnInquiryBtnStatus(DeviceHandle)
                         Tester.Print i; Space(5); "Key press value="; ReturnValue
                         i = i + 1
                     Loop While i < 3 And ReturnValue <> 10
                    ' fnFreeDeviceHandle (DeviceHandle)
                   '  fnFreeDeviceHandle (DeviceHandle)
                     
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
                Tester.Print rv5, " \\CF overcurrent protect"
                Tester.Print rv6, " \\CF power , 6: over upper limit, 5: under lower limit"
                Tester.Print "LBA="; LBA
                
AU6377ALFResult:
                       
                        
                       If rv6 = 5 Or rv6 = 6 Or rv5 <> 1 Then
                        
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                            
                            If rv0 * rv1 * rv2 * rv3 * rv4 = 1 Then
                              XDWriteFail = XDWriteFail + 1
                              TestResult = "XD_WF"
                            End If
                            
                            
                        ElseIf rv0 = UNKNOW Then
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
                        ElseIf rv2 = WRITE_FAIL Or rv3 = WRITE_FAIL Or rv4 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                             CFWriteFail = CFWriteFail + 1
                            TestResult = "CF_WF"
                        ElseIf rv2 = READ_FAIL Or rv3 = READ_FAIL Or rv4 = READ_FAIL Or rv5 = READ_FAIL Then
                            CFWriteFail = CFWriteFail + 1
                            TestResult = "CF_WF"
        
                        ElseIf rv5 * rv6 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        
                        End If
                
               
                  
              
                
                
                
     '    CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off
     '    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
     '    Call MsecDelay(0.8)
     '    CardResult = DO_WritePort(card, Channel_P1A, &H40)
         
                
                
               
   End Sub
   
Public Sub AU6476JLO10TestSub()
'AU6476BLF24 add MS bus width test, and modify SD bus width
 
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
               
   
TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
Tester.Print "AU6476JL is NB mode"
                  
 If Left(TmpChip, 8) = "AU6378AL" Or Left(TmpChip, 8) = "AU6476JL" Then
 
         CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  // force into reader mode
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.2)
         CardResult = DO_WritePort(card, Channel_P1A, &H0)
         
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H3F)
          
          
  End If
  
              
  
  
  '========================== AU6377 new board switch assign ment  ============
If TmpChip = "AU6377ALF21" Then ' this for new board  and internalrom
         RomSelector = &H10  '-------- this is for MS in pin
  End If
  
  
  If TmpChip = "AU6377ALF24" Then ' this for new board  and internalrom
         RomSelector = &H10
  End If
  
  If TmpChip = "AU6377ALF25" Then ' this for new board  and internalrom
         RomSelector = &H0
  End If
         
         
  If Left(TmpChip, 10) = "AU6377ALF2" Then
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector)  ' 1111 1110
         CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' 1111 1110
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector) ' 5th bit is rom selector, High is internal rom
  End If
  
  


  
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1.3)
               
                
                 If GetDeviceName("vid") <> "" Then
                    rv0 = 0
                    Tester.Print "NB mode test Fail"
                    GoTo AU6377ALFResult
                  
                  End If
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
                Tester.Print LBA
                If TmpChip = "AU6377ALF25" Or TmpChip = "AU6378ALF21" Then
                  VidName = "vid_1984"
                Else
                 VidName = "vid_058f"
                End If
                
              
              
              
              
              
              
             
 '================================= Test light off =============================
                
              
                
                
                   
                    CardResult = DO_WritePort(card, Channel_P1A, &H18)      ' &H8 overcurrent, &H18 no overcurrent
                  
                    
                   Call MsecDelay(1.3)
                 
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
                 rv0 = CBWTest_New(0, 1, VidName)    ' sd slot
                 
                 
                 If rv0 = 1 Then
                    rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                    If rv0 <> 1 Then
                    Tester.Print "SD bus width Fail"
                    rv0 = 2
                    End If
                 End If
                 
                 
                 
                 If rv0 = 1 And TmpChip = "AU6375HLF22" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                End If
                
                   If Left(TmpChip, 10) = "AU6377ALF2" Then
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                      LBA = TmpLBA
                   End If
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                  rv1 = CBWTest_NewCFOverCurrentRWnoPower(1, rv0, VidName)    ' cf slot
                  
                  If rv1 = 5 Then
                     Tester.Print "OverCurrent wrong action"
                  End If
                  ' rv1 = 1
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
              
             
                 rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
               
                ClosePipe
                
                 If rv1 = 1 And rv2 <> 1 Then
                 Tester.Label9.Caption = "SMC FAIL"
                End If
                '============= SMC test begin =======================================
               
              
                 
                          '--- for SMC
                CardResult = DO_WritePort(card, Channel_P1A, &H1C)   ' 0110 0100
                Call MsecDelay(0.1)
                      
                
                CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100 SMC high
                Call MsecDelay(0.1)
                ClosePipe
                
                
               
                  rv3 = CBWTest_New(2, rv2, VidName)
               
                Call LabelMenu(2, rv3, rv2)
                ClosePipe
                  If rv2 = 1 And rv3 <> 1 Then
                 Tester.Label9.Caption = "XD FAIL"
                End If
                
                
              
                
               '=============== SMC test END ==================================================
               
              
               rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
               
               If rv4 = 1 Then
               rv4 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                 If rv4 <> 1 Then
                   Tester.Print "MS bus width Fail"
                   rv4 = 2
                 End If
              End If
            
               ClosePipe
               Call LabelMenu(3, rv4, rv3)
               
 '======================== light test ======================
               
           If rv4 = 1 Then
               CardResult = DO_WritePort(card, Channel_P1A, &H6)   ' 0110 0100 remove CF, and do overcurrent
               Call MsecDelay(0.5)
               CardResult = DO_WritePort(card, Channel_P1A, &H4)     ' CF card in
                 
                    
               Call MsecDelay(0.5)
               
               
                rv4 = CBWTest_NewOverCurrent(1, 1, VidName)     ' Test CF at 1st slot
                 
       
                  
                     
                    If rv4 = 3 Then
                    rv4 = 1
                    Else
                    rv4 = 5
                    Tester.Print "over current protect fail"
                    End If
               
          End If
               
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
      If rv4 = 1 Then
          CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off  for HID mode
          result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
          Call MsecDelay(0.4)
          
          CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
                            
    
          Call MsecDelay(1.5)
          CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
          If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
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
                         GoTo AU6377ALFResult
                     End If
                     
                     '=======================
                     '  key press test, it will return 8 when key up, GPI 6 must do low go hi action
                     '========================
                     
                
                     Do
                          CardResult = DO_WritePort(card, Channel_P1A, &H40) 'GPI6 : bit 6: pull high
                          Sleep (200)
                         CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' GPI6 : bit 6: pull low
                         Sleep (1000)
                       
                         ReturnValue = fnInquiryBtnStatus(DeviceHandle)
                         Tester.Print i; Space(5); "Key press value="; ReturnValue
                         i = i + 1
                     Loop While i < 3 And ReturnValue <> 10
                    ' fnFreeDeviceHandle (DeviceHandle)
                   '  fnFreeDeviceHandle (DeviceHandle)
                     
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
                
AU6377ALFResult:
                        If rv1 = 5 Or rv4 = 5 Then
                        
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        
                        ElseIf rv0 = UNKNOW Then
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
                        ElseIf rv2 = WRITE_FAIL Or rv3 = WRITE_FAIL Or rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv2 = READ_FAIL Or rv3 = READ_FAIL Or rv4 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "MS_RF"
        
                        ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        
                        End If
                
               
                  
              
                
                
                
     '    CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off
     '    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
     '    Call MsecDelay(0.8)
     '    CardResult = DO_WritePort(card, Channel_P1A, &H40)
         
                
                
               
   End Sub
   
 Public Sub AU6476BLF22TestSub()
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
               
   
TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
Tester.Print "AU6476BL is NB mode"
                  
 If Left(TmpChip, 8) = "AU6378AL" Or Left(TmpChip, 8) = "AU6476BL" Then
 
         CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  // force into reader mode
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.2)
         CardResult = DO_WritePort(card, Channel_P1A, &H0)
         
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H3F)
          
          
  End If
  
              
  
  
  '========================== AU6377 new board switch assign ment  ============
If TmpChip = "AU6377ALF21" Then ' this for new board  and internalrom
         RomSelector = &H10  '-------- this is for MS in pin
  End If
  
  
  If TmpChip = "AU6377ALF24" Then ' this for new board  and internalrom
         RomSelector = &H10
  End If
  
  If TmpChip = "AU6377ALF25" Then ' this for new board  and internalrom
         RomSelector = &H0
  End If
         
         
  If Left(TmpChip, 10) = "AU6377ALF2" Then
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector)  ' 1111 1110
         CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' 1111 1110
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector) ' 5th bit is rom selector, High is internal rom
  End If
  
  


  
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1.3)
               
                
                 If GetDeviceName("vid") <> "" Then
                    rv0 = 0
                    Tester.Print "NB mode test Fail"
                    GoTo AU6377ALFResult
                  
                  End If
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
                Tester.Print LBA
                If TmpChip = "AU6377ALF25" Or TmpChip = "AU6378ALF21" Then
                  VidName = "vid_1984"
                Else
                 VidName = "vid_058f"
                End If
                
              
             
 '================================= Test light off =============================
                
              
                
                
                   
                    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                  
                    
                   Call MsecDelay(1.3)
                 
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
                 rv0 = CBWTest_New(0, 1, VidName)     ' sd slot
                 
                 If rv0 = 1 And TmpChip = "AU6375HLF22" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                End If
                
                   If Left(TmpChip, 10) = "AU6377ALF2" Then
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                      LBA = TmpLBA
                   End If
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                  rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
                  ' rv1 = 1
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
              
             
                 rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
               
                ClosePipe
                
                 If rv1 = 1 And rv2 <> 1 Then
                 Tester.Label9.Caption = "SMC FAIL"
                End If
                '============= SMC test begin =======================================
               
              
                 
                          '--- for SMC
                CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
                Call MsecDelay(0.1)
                      
                
                CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
                Call MsecDelay(0.1)
                ClosePipe
                
                
               
                  rv3 = CBWTest_New(2, rv2, VidName)
               
                Call LabelMenu(2, rv3, rv2)
                ClosePipe
                  If rv2 = 1 And rv3 <> 1 Then
                 Tester.Label9.Caption = "XD FAIL"
                End If
                
                
              
                
               '=============== SMC test END ==================================================
               
              
               rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
            
               ClosePipe
               Call LabelMenu(3, rv4, rv3)
               
 '======================== light test ======================
               
               
               
               
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
      If rv4 = 1 Then
          CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off  for HID mode
          result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
          Call MsecDelay(0.4)
          
          CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
                            
    
          Call MsecDelay(1.2)
          CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
          If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
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
                         GoTo AU6377ALFResult
                     End If
                     
                     '=======================
                     '  key press test, it will return 8 when key up, GPI 6 must do low go hi action
                     '========================
                     
                
                     Do
                          CardResult = DO_WritePort(card, Channel_P1A, &H40) 'GPI6 : bit 6: pull high
                          Sleep (200)
                         CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' GPI6 : bit 6: pull low
                         Sleep (1000)
                       
                         ReturnValue = fnInquiryBtnStatus(DeviceHandle)
                         Tester.Print i; Space(5); "Key press value="; ReturnValue
                         i = i + 1
                     Loop While i < 3 And ReturnValue <> 10
                    ' fnFreeDeviceHandle (DeviceHandle)
                   '  fnFreeDeviceHandle (DeviceHandle)
                     
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
                
AU6377ALFResult:
                        
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
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        
                        End If
                
               
                  
              
                
                
                
     '    CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off
     '    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
     '    Call MsecDelay(0.8)
     '    CardResult = DO_WritePort(card, Channel_P1A, &H40)
         
                
                
               
   End Sub
   

Public Sub AU6476CLF21TestSub()
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
               
   
TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
  
                  
 
 
         CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  // force into reader mode
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.2)
         CardResult = DO_WritePort(card, Channel_P1A, &H0)
         
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H3F)
          
          
   
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1.6)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
               ' Dim ChipString As String
                 VidName = "vid_058f"
                
                        
                 
              
                ClosePipe
                 rv0 = CBWTest_New_no_card(0, 1, VidName)
                'Tester.print "a1"
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                rv1 = CBWTest_New_no_card(1, rv0, VidName)
               '  Tester.print "a2"
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
               
               rv2 = CBWTest_New_no_card(2, rv1, VidName)
               '  Tester.print "a3"
                 Call LabelMenu(2, rv2, rv1)
                ClosePipe
                  rv3 = CBWTest_New_no_card(3, rv2, VidName)
                ' Tester.print "a4"
                ClosePipe
               Call LabelMenu(3, rv3, rv2)
                
 '================================= Test light off =============================
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv3, " \\MS :0 Unknow device, 1 pass ,2 card change bit fail"
                 
                
                
             
                
                ' test chip
                      '    CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If LightOff <> 255 Then
                            UsbSpeedTestResult = GPO_FAIL
                             rv0 = 2
                         End If
          
               Tester.Print rv0, " \\LED Fail"
                
                
                
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
               
                       
       
                 
              
'====================================== Assing R/W test switch =====================================
                 
                If TestResult = "PASS" Then
                   TestResult = ""
                   
                    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                   
              
                    
                   Call MsecDelay(0.1)
                 
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
                 rv0 = CBWTest_New(0, 1, VidName)     ' sd slot
                 
                 If rv0 = 1 And TmpChip = "AU6375HLF22" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                End If
                
                   If Left(TmpChip, 10) = "AU6377ALF2" Then
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                      LBA = TmpLBA
                   End If
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
                
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
              
                
                 rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
               
                ClosePipe
                
                 If rv1 = 1 And rv2 <> 1 Then
                 Tester.Label9.Caption = "SMC FAIL"
                End If
                '============= SMC test begin =======================================
               
              
                 
                          '--- for SMC
                CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
                Call MsecDelay(0.1)
                      
                
                CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
                Call MsecDelay(0.1)
                ClosePipe
                
                
               
                  rv3 = CBWTest_New(2, rv2, VidName)
               
                Call LabelMenu(2, rv3, rv2)
                ClosePipe
                  If rv2 = 1 And rv3 <> 1 Then
                 Tester.Label9.Caption = "XD FAIL"
                End If
                
                
              
                
               '=============== SMC test END ==================================================
               
              
               rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
            
               ClosePipe
               Call LabelMenu(3, rv4, rv3)
               
 '======================== light test ======================
            Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            Tester.Print "LBA="; LBA
               
               
               
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
      If rv4 = 1 Then
          CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off  for HID mode
          result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
          Call MsecDelay(0.4)
          
          CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
                            
    
          Call MsecDelay(1.2)
          CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
          If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
                  UsbSpeedTestResult = GPO_FAIL
                    rv4 = 2
                    Tester.Label9.Caption = "GPO FAIL " & LightOff
           End If
     End If
           Tester.Print rv4, " \\GPO,LED fail"
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
                         GoTo AU6377ALFResult
                     End If
                     
                     '=======================
                     '  key press test, it will return 8 when key up, GPI 6 must do low go hi action
                     '========================
                     
                
                     Do
                          CardResult = DO_WritePort(card, Channel_P1A, &H40) 'GPI6 : bit 6: pull high
                          Sleep (200)
                         CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' GPI6 : bit 6: pull low
                         Sleep (1000)
                       
                         ReturnValue = fnInquiryBtnStatus(DeviceHandle)
                         Tester.Print i; Space(5); "Key press value="; ReturnValue
                         i = i + 1
                     Loop While i < 3 And ReturnValue <> 10
                    ' fnFreeDeviceHandle (DeviceHandle)
                   '  fnFreeDeviceHandle (DeviceHandle)
                     
                     If ReturnValue <> 10 Then
                     
                      rv1 = 2
                     
                       Call LabelMenu(1, rv1, rv0)
                       Label9.Caption = "KeyPress Fail"
                       
                     End If
                              
                 
           End If
                
                
                Tester.Print rv0, " \\Can not Find HID deevice"
                Tester.Print rv1, " \\Key press Fail"
             '   Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
             '   Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
             '   Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
             '   Tester.Print "LBA="; LBA
                
AU6377ALFResult:
                        
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
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        
                        End If
                
               
                  
                End If
                
                
                
     '    CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off
     '    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
     '    Call MsecDelay(0.8)
     '    CardResult = DO_WritePort(card, Channel_P1A, &H40)
         
                
                
               
   End Sub
   
 Public Sub AU6476CLF22TestSub()
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
               
   
TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
  
                  
 
 
         CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  // force into reader mode
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.2)
         CardResult = DO_WritePort(card, Channel_P1A, &H0)
         
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H3F)
          
          
   
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1.6)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
               ' Dim ChipString As String
                 VidName = "vid_058f"
                
                        
                 
              
                ClosePipe
                 rv0 = CBWTest_New_no_card(0, 1, VidName)
                'Tester.print "a1"
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                rv1 = CBWTest_New_no_card(1, rv0, VidName)
               '  Tester.print "a2"
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
               
               rv2 = CBWTest_New_no_card(2, rv1, VidName)
               '  Tester.print "a3"
                 Call LabelMenu(2, rv2, rv1)
                ClosePipe
                  rv3 = CBWTest_New_no_card(3, rv2, VidName)
                ' Tester.print "a4"
                ClosePipe
               Call LabelMenu(3, rv3, rv2)
                
 '================================= Test light off =============================
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv3, " \\MS :0 Unknow device, 1 pass ,2 card change bit fail"
                 
                
                
             
                
                ' test chip
                      '    CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If LightOff <> 255 Then
                            UsbSpeedTestResult = GPO_FAIL
                             rv0 = 2
                         End If
          
               Tester.Print rv0, " \\LED Fail"
                
                
                
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
               
                       
       
                 
              
'====================================== Assing R/W test switch =====================================
                 
                If TestResult = "PASS" Then
                   TestResult = ""
                   
                    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                   
              
                    
                   Call MsecDelay(0.1)
                 
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
                 rv0 = CBWTest_New_SD_Speed(0, 1, VidName, "4Bits")    ' sd slot
                 
                 If rv0 = 1 And TmpChip = "AU6375HLF22" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                End If
                
                   If Left(TmpChip, 10) = "AU6377ALF2" Then
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                      LBA = TmpLBA
                   End If
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
                
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
              
                
                 rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
               
                ClosePipe
                
                 If rv1 = 1 And rv2 <> 1 Then
                 Tester.Label9.Caption = "SMC FAIL"
                End If
                '============= SMC test begin =======================================
               
              
                 
                          '--- for SMC
                CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
                Call MsecDelay(0.1)
                      
                
                CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
                Call MsecDelay(0.1)
                ClosePipe
                
                
               
                  rv3 = CBWTest_New(2, rv2, VidName)
               
                Call LabelMenu(2, rv3, rv2)
                ClosePipe
                  If rv2 = 1 And rv3 <> 1 Then
                 Tester.Label9.Caption = "XD FAIL"
                End If
                
                
              
                
               '=============== SMC test END ==================================================
               
              
               rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
            
               ClosePipe
               Call LabelMenu(3, rv4, rv3)
               
 '======================== light test ======================
            Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            Tester.Print "LBA="; LBA
               
               
               
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
      If rv4 = 1 Then
          CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off  for HID mode
          result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
          Call MsecDelay(0.4)
          
          CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
                            
    
          Call MsecDelay(1.2)
          CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
          If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
                  UsbSpeedTestResult = GPO_FAIL
                    rv4 = 2
                    Tester.Label9.Caption = "GPO FAIL " & LightOff
           End If
     End If
           Tester.Print rv4, " \\GPO,LED fail"
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
                         GoTo AU6377ALFResult
                     End If
                     
                     '=======================
                     '  key press test, it will return 8 when key up, GPI 6 must do low go hi action
                     '========================
                     
                
                     Do
                          CardResult = DO_WritePort(card, Channel_P1A, &H40) 'GPI6 : bit 6: pull high
                          Sleep (200)
                         CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' GPI6 : bit 6: pull low
                         Sleep (1000)
                       
                         ReturnValue = fnInquiryBtnStatus(DeviceHandle)
                         Tester.Print i; Space(5); "Key press value="; ReturnValue
                         i = i + 1
                     Loop While i < 3 And ReturnValue <> 10
                    ' fnFreeDeviceHandle (DeviceHandle)
                   '  fnFreeDeviceHandle (DeviceHandle)
                     
                     If ReturnValue <> 10 Then
                     
                      rv1 = 2
                     
                       Call LabelMenu(1, rv1, rv0)
                       Label9.Caption = "KeyPress Fail"
                       
                     End If
                              
                 
           End If
                
                
                Tester.Print rv0, " \\Can not Find HID deevice"
                Tester.Print rv1, " \\Key press Fail"
             '   Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
             '   Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
             '   Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
             '   Tester.Print "LBA="; LBA
                
AU6377ALFResult:
                        
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
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        
                        End If
                
               
                  
                End If
                
                
                
     '    CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off
     '    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
     '    Call MsecDelay(0.8)
     '    CardResult = DO_WritePort(card, Channel_P1A, &H40)
         
                
                
               
   End Sub
   
 Public Sub AU6476MLF24TestSub()
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
               
   
TmpChip = ChipName
'==================================== Switch assign ==========================================
ChipName = "AU6476CLF24"
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
  
                  
 
 
         CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  // force into reader mode
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.2)
         CardResult = DO_WritePort(card, Channel_P1A, &H0)
         
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H3F)
          
          
   
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1.6)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
               ' Dim ChipString As String
                 VidName = "vid_058f"
                
                        
                 
              
                ClosePipe
                 rv0 = CBWTest_New_no_card(0, 1, VidName)
                'Tester.print "a1"
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                rv1 = CBWTest_New_no_card(1, rv0, VidName)
               '  Tester.print "a2"
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
               
               rv2 = CBWTest_New_no_card(2, rv1, VidName)
               '  Tester.print "a3"
                 Call LabelMenu(2, rv2, rv1)
                ClosePipe
                  rv3 = CBWTest_New_no_card(3, rv2, VidName)
                ' Tester.print "a4"
                ClosePipe
               Call LabelMenu(3, rv3, rv2)
                
 '================================= Test light off =============================
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv3, " \\MS :0 Unknow device, 1 pass ,2 card change bit fail"
                 
                
                
             
                
                ' test chip
                      '    CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If LightOff <> 255 Then
                            UsbSpeedTestResult = GPO_FAIL
                             rv0 = 2
                         End If
          
               Tester.Print rv0, " \\LED Fail"
                
                
                
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
               
                       
       
                 
              
'====================================== Assing R/W test switch =====================================
                 
                If TestResult = "PASS" Then
                   TestResult = ""
                   
                    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                   
              
                    
                   Call MsecDelay(0.1)
                 
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
                 rv0 = CBWTest_New(0, 1, VidName)    ' sd slot
                 
                 If rv0 = 1 Then
                    rv0 = Read_SD_Speed(0, 0, 64, "4Bits")
                    If rv0 <> 1 Then
                      Tester.Print "SD bus width Fail"
                      rv0 = 2
                    End If
                 End If
                   
                 
                 If rv0 = 1 And TmpChip = "AU6375HLF22" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                End If
                
                   If Left(TmpChip, 10) = "AU6377ALF2" Then
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                      LBA = TmpLBA
                   End If
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
                
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
              
                
                 rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
               
                ClosePipe
                
                 If rv1 = 1 And rv2 <> 1 Then
                 Tester.Label9.Caption = "SMC FAIL"
                End If
                '============= SMC test begin =======================================
               
              
                 
                          '--- for SMC
                CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
                Call MsecDelay(0.1)
                      
                
                CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
                Call MsecDelay(0.1)
                ClosePipe
                
                
               
                  rv3 = CBWTest_New(2, rv2, VidName)
               
                Call LabelMenu(2, rv3, rv2)
                ClosePipe
                  If rv2 = 1 And rv3 <> 1 Then
                 Tester.Label9.Caption = "XD FAIL"
                End If
                
                
              
                
               '=============== SMC test END ==================================================
               
              
               rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
               
               If rv4 = 1 Then
                  rv4 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                  If rv4 <> 1 Then
                    rv4 = 2
                    Tester.Print "MS bus width Fail"
                    End If
               End If
                    
               
               
            
               ClosePipe
               Call LabelMenu(3, rv4, rv3)
               
               
               
 '======================== light test ======================
            Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            Tester.Print "LBA="; LBA
               
               
               
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
      If rv4 = 1 Then
          CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off  for HID mode
          result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
          Call MsecDelay(0.4)
          
          CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
                            
    
          Call MsecDelay(1.2)
          CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
          If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
                  UsbSpeedTestResult = GPO_FAIL
                    rv4 = 2
                    Tester.Label9.Caption = "GPO FAIL " & LightOff
           End If
     End If
           Tester.Print rv4, " \\GPO,LED fail"
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
                         GoTo AU6377ALFResult
                     End If
                     
                     '=======================
                     '  key press test, it will return 8 when key up, GPI 6 must do low go hi action
                     '========================
                     
                
                     Do
                          CardResult = DO_WritePort(card, Channel_P1A, &H40) 'GPI6 : bit 6: pull high
                          Sleep (200)
                         CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' GPI6 : bit 6: pull low
                         Sleep (1000)
                       
                         ReturnValue = fnInquiryBtnStatus(DeviceHandle)
                         Tester.Print i; Space(5); "Key press value="; ReturnValue
                         i = i + 1
                     Loop While i < 3 And ReturnValue <> 10
                    ' fnFreeDeviceHandle (DeviceHandle)
                   '  fnFreeDeviceHandle (DeviceHandle)
                     
                     If ReturnValue <> 10 Then
                     
                      rv1 = 2
                     
                       Call LabelMenu(1, rv1, rv0)
                       Label9.Caption = "KeyPress Fail"
                       
                     End If
                              
                 
           End If
                
                
                Tester.Print rv0, " \\Can not Find HID deevice"
                Tester.Print rv1, " \\Key press Fail"
             '   Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
             '   Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
             '   Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
             '   Tester.Print "LBA="; LBA
                
AU6377ALFResult:
                        
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
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        
                        End If
                
               
                  
                End If
                
                
                
     '    CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off
     '    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
     '    Call MsecDelay(0.8)
     '    CardResult = DO_WritePort(card, Channel_P1A, &H40)
         
                
                
               
   End Sub
   
 Public Sub AU6476CLF24TestSub()
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
               
   
TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
  
                  
 
 
         CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  // force into reader mode
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.2)
         CardResult = DO_WritePort(card, Channel_P1A, &H0)
         
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H3F)
          
          
   
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1.6)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
               ' Dim ChipString As String
                 VidName = "vid_058f"
                
                        
                 
              
                ClosePipe
                 rv0 = CBWTest_New_no_card(0, 1, VidName)
                'Tester.print "a1"
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                rv1 = CBWTest_New_no_card(1, rv0, VidName)
               '  Tester.print "a2"
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
               
               rv2 = CBWTest_New_no_card(2, rv1, VidName)
               '  Tester.print "a3"
                 Call LabelMenu(2, rv2, rv1)
                ClosePipe
                  rv3 = CBWTest_New_no_card(3, rv2, VidName)
                ' Tester.print "a4"
                ClosePipe
               Call LabelMenu(3, rv3, rv2)
                
 '================================= Test light off =============================
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv3, " \\MS :0 Unknow device, 1 pass ,2 card change bit fail"
                 
                
                
             
                
                ' test chip
                      '    CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If LightOff <> 255 Then
                            UsbSpeedTestResult = GPO_FAIL
                             rv0 = 2
                         End If
          
               Tester.Print rv0, " \\LED Fail"
                
                
                
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
               
                       
       
                 
              
'====================================== Assing R/W test switch =====================================
                 
                If TestResult = "PASS" Then
                   TestResult = ""
                   
                    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                   
              
                    
                   Call MsecDelay(0.1)
                 
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
                 rv0 = CBWTest_New(0, 1, VidName)    ' sd slot
                 
                 If rv0 = 1 Then
                    rv0 = Read_SD_Speed(0, 0, 64, "4Bits")
                    If rv0 <> 1 Then
                      Tester.Print "SD bus width Fail"
                      rv0 = 2
                    End If
                 End If
                   
                 
                 If rv0 = 1 And TmpChip = "AU6375HLF22" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                End If
                
                   If Left(TmpChip, 10) = "AU6377ALF2" Then
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                      LBA = TmpLBA
                   End If
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
                
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
              
                
                 rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
               
                ClosePipe
                
                 If rv1 = 1 And rv2 <> 1 Then
                 Tester.Label9.Caption = "SMC FAIL"
                End If
                '============= SMC test begin =======================================
               
              
                 
                          '--- for SMC
                CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
                Call MsecDelay(0.1)
                      
                
                CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
                Call MsecDelay(0.1)
                ClosePipe
                
                
               
                  rv3 = CBWTest_New(2, rv2, VidName)
               
                Call LabelMenu(2, rv3, rv2)
                ClosePipe
                  If rv2 = 1 And rv3 <> 1 Then
                 Tester.Label9.Caption = "XD FAIL"
                End If
                
                
              
                
               '=============== SMC test END ==================================================
               
              
               rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
               
               If rv4 = 1 Then
                  rv4 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                  If rv4 <> 1 Then
                    rv4 = 2
                    Tester.Print "MS bus width Fail"
                    End If
               End If
                    
               
               
            
               ClosePipe
               Call LabelMenu(3, rv4, rv3)
               
               
               
 '======================== light test ======================
            Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            Tester.Print "LBA="; LBA
               
               
               
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
      If rv4 = 1 Then
          CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off  for HID mode
          result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
          Call MsecDelay(0.4)
          
          CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
                            
    
          Call MsecDelay(1.5)
          CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
          If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
                  UsbSpeedTestResult = GPO_FAIL
                    rv4 = 2
                    Tester.Label9.Caption = "GPO FAIL " & LightOff
           End If
     End If
           Tester.Print rv4, " \\GPO,LED fail"
        If rv4 = 1 Then
        ' code begin
                     
                     Tester.Cls
                     Tester.Print "keypress test begin---------------"
                     Dim ReturnValue As Byte
                     
                     DeviceHandle = &HFFFF  'invalid handle initial value
                     Call MsecDelay(0.1)
                     ReturnValue = fnGetDeviceHandle(DeviceHandle)
                     Tester.Print ReturnValue; Space(5); ' 1: pass the other refer btnstatus.h
                     Tester.Print "DeviceHandle="; DevicehHandle
                     
                     If ReturnValue <> 1 Then
                           rv0 = UNKNOW       '---> HID mode unknow device mode
                           Call LabelMenu(0, rv0, 1)
                           Tester.Label9.Caption = "HID mode unknow device"
                          fnFreeDeviceHandle (DeviceHandle)
                         GoTo AU6377ALFResult
                     End If
                     
                     '=======================
                     '  key press test, it will return 8 when key up, GPI 6 must do low go hi action
                     '========================
                     
                
                     Do
                          CardResult = DO_WritePort(card, Channel_P1A, &H40) 'GPI6 : bit 6: pull high
                          Sleep (200)
                         CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' GPI6 : bit 6: pull low
                         Sleep (1000)
                       
                         ReturnValue = fnInquiryBtnStatus(DeviceHandle)
                         Tester.Print i; Space(5); "Key press value="; ReturnValue
                         i = i + 1
                     Loop While i < 3 And ReturnValue <> 10
                    ' fnFreeDeviceHandle (DeviceHandle)
                   '  fnFreeDeviceHandle (DeviceHandle)
                     
                     If ReturnValue <> 10 Then
                     
                      rv1 = 2
                     
                       Call LabelMenu(1, rv1, rv0)
                       Label9.Caption = "KeyPress Fail"
                       
                     End If
                              
                 
           End If
                
                
                Tester.Print rv0, " \\Can not Find HID deevice"
                Tester.Print rv1, " \\Key press Fail"
             '   Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
             '   Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
             '   Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
             '   Tester.Print "LBA="; LBA
                
AU6377ALFResult:
                        
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
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        
                        End If
                
               
                  
                End If
                
                
                
     '    CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off
     '    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
     '    Call MsecDelay(0.8)
     '    CardResult = DO_WritePort(card, Channel_P1A, &H40)
         
                
                
               
   End Sub
    Public Sub AU6476CLF25TestSub()
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
               
   
TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
  
                  
 
 
         CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  // force into reader mode
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.2)
         CardResult = DO_WritePort(card, Channel_P1A, &H0)
         
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H3F)
          
          
   
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1.6)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
               ' Dim ChipString As String
                 VidName = "vid_058f"
                
                        
                 
              
                ClosePipe
                 rv0 = CBWTest_New_no_card(0, 1, VidName)
                'Tester.print "a1"
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                rv1 = CBWTest_New_no_card(1, rv0, VidName)
               '  Tester.print "a2"
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
               
               rv2 = CBWTest_New_no_card(2, rv1, VidName)
               '  Tester.print "a3"
                 Call LabelMenu(2, rv2, rv1)
                ClosePipe
                  rv3 = CBWTest_New_no_card(3, rv2, VidName)
                ' Tester.print "a4"
                ClosePipe
               Call LabelMenu(3, rv3, rv2)
                
 '================================= Test light off =============================
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv3, " \\MS :0 Unknow device, 1 pass ,2 card change bit fail"
                 
                
                
             
                
                ' test chip
                      '    CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If LightOff <> 255 Then
                            UsbSpeedTestResult = GPO_FAIL
                             rv0 = 2
                         End If
          
               Tester.Print rv0, " \\LED Fail"
                
                
                
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
               
                       
       
                 
              
'====================================== Assing R/W test switch =====================================
                 
                If TestResult = "PASS" Then
                   TestResult = ""
                   
                    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                   
              
                    
                   Call MsecDelay(0.1)
                 
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
                 rv0 = CBWTest_New(0, 1, VidName)    ' sd slot
                 
                 If rv0 = 1 Then
                    rv0 = Read_SD_Speed(0, 0, 64, "4Bits")
                    If rv0 <> 1 Then
                      Tester.Print "SD bus width Fail"
                      rv0 = 2
                    End If
                 End If
                   
                 
                 If rv0 = 1 And TmpChip = "AU6375HLF22" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                End If
                
                   If Left(TmpChip, 10) = "AU6377ALF2" Then
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                      LBA = TmpLBA
                   End If
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
                
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
                
                rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
               
                ClosePipe
                
                 If rv1 = 1 And rv2 <> 1 Then
                 Tester.Label9.Caption = "SMC FAIL"
                End If
                '============= SMC test begin =======================================
               
              
                 
                          '--- for SMC
                CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
                Call MsecDelay(0.1)
                      
                
                CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
                Call MsecDelay(0.1)
                ClosePipe
                
                
               
                  rv3 = CBWTest_New(2, rv2, VidName)
               
                Call LabelMenu(2, rv3, rv2)
                ClosePipe
                  If rv2 = 1 And rv3 <> 1 Then
                 Tester.Label9.Caption = "XD FAIL"
                End If
                
                
              
                
               '=============== SMC test END ==================================================
               
              
               rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
               
               If rv4 = 1 Then
                  rv4 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                  If rv4 <> 1 Then
                    rv4 = 2
                    Tester.Print "MS bus width Fail"
                    End If
               End If
                    
               
               
            
               ClosePipe
               Call LabelMenu(3, rv4, rv3)
               
               
               
 '======================== light test ======================
            Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            Tester.Print "LBA="; LBA
               
               
               
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
      If rv4 = 1 Then
          CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off  for HID mode
          result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
          Call MsecDelay(0.4)
          
          CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
                            
    
          Call MsecDelay(1.5)
          CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
          If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
                  UsbSpeedTestResult = GPO_FAIL
                    rv4 = 2
                    Tester.Label9.Caption = "GPO FAIL " & LightOff
           End If
     End If
           Tester.Print rv4, " \\GPO,LED fail"
        If rv4 = 1 Then
        ' code begin
                     
                     Tester.Cls
                     Tester.Print "keypress test begin---------------"
                     Dim ReturnValue As Byte
                     
                     DeviceHandle = &HFFFF  'invalid handle initial value
                     Call MsecDelay(0.1)
                     ReturnValue = fnGetDeviceHandle(DeviceHandle)
                     Tester.Print ReturnValue; Space(5); ' 1: pass the other refer btnstatus.h
                     Tester.Print "DeviceHandle="; DevicehHandle
                     
                     If ReturnValue <> 1 Then
                           rv0 = UNKNOW       '---> HID mode unknow device mode
                           Call LabelMenu(0, rv0, 1)
                           Tester.Label9.Caption = "HID mode unknow device"
                          fnFreeDeviceHandle (DeviceHandle)
                         GoTo AU6377ALFResult
                     End If
                     
                     '=======================
                     '  key press test, it will return 8 when key up, GPI 6 must do low go hi action
                     '========================
                     
                
                     Do
                          CardResult = DO_WritePort(card, Channel_P1A, &H40) 'GPI6 : bit 6: pull high
                          Sleep (200)
                         CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' GPI6 : bit 6: pull low
                         Sleep (1000)
                       
                         ReturnValue = fnInquiryBtnStatus(DeviceHandle)
                         Tester.Print i; Space(5); "Key press value="; ReturnValue
                         i = i + 1
                     Loop While i < 3 And ReturnValue <> 11
                    ' fnFreeDeviceHandle (DeviceHandle)
                   '  fnFreeDeviceHandle (DeviceHandle)
                     
                     If ReturnValue <> 11 Then
                     
                      rv1 = 2
                     
                       Call LabelMenu(1, rv1, rv0)
                       Label9.Caption = "KeyPress Fail"
                       
                     End If
                              
                 
           End If
                
                
                Tester.Print rv0, " \\Can not Find HID deevice"
                Tester.Print rv1, " \\Key press Fail"
             '   Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
             '   Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
             '   Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
             '   Tester.Print "LBA="; LBA
                
AU6377ALFResult:
                        
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
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        
                        End If
                
               
                  
                End If
                
                
                
     '    CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off
     '    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
     '    Call MsecDelay(0.8)
     '    CardResult = DO_WritePort(card, Channel_P1A, &H40)
         
                
                
               
   End Sub
Public Sub AU6476CLF26TestSub()
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
               
'2010/9/3  copy from AU6476CLF24 purpose to solve HID mode unstable issue
   
TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
  
                  
 
 
         CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  // force into reader mode
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.2)
         CardResult = DO_WritePort(card, Channel_P1A, &H0)
         
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H3F)
          
          
   
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1.6)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
               ' Dim ChipString As String
                 VidName = "vid_058f"
                
                        
                 
              
                ClosePipe
                 rv0 = CBWTest_New_no_card(0, 1, VidName)
                'Tester.print "a1"
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                rv1 = CBWTest_New_no_card(1, rv0, VidName)
               '  Tester.print "a2"
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
               
               rv2 = CBWTest_New_no_card(2, rv1, VidName)
               '  Tester.print "a3"
                 Call LabelMenu(2, rv2, rv1)
                ClosePipe
                  rv3 = CBWTest_New_no_card(3, rv2, VidName)
                ' Tester.print "a4"
                ClosePipe
               Call LabelMenu(3, rv3, rv2)
                
 '================================= Test light off =============================
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv3, " \\MS :0 Unknow device, 1 pass ,2 card change bit fail"
                 
                
                
             
                
                ' test chip
                      '    CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If LightOff <> 255 Then
                            UsbSpeedTestResult = GPO_FAIL
                             rv0 = 2
                         End If
          
               Tester.Print rv0, " \\LED Fail"
                
                
                
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
               
                       
       
                 
              
'====================================== Assing R/W test switch =====================================
                 
                If TestResult = "PASS" Then
                   TestResult = ""
                   
                    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                   
              
                    
                   Call MsecDelay(0.1)
                 
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
                 rv0 = CBWTest_New(0, 1, VidName)    ' sd slot
                 
                 If rv0 = 1 Then
                    rv0 = Read_SD_Speed(0, 0, 64, "4Bits")
                    If rv0 <> 1 Then
                      Tester.Print "SD bus width Fail"
                      rv0 = 2
                    End If
                 End If
                   
                 
                 If rv0 = 1 And TmpChip = "AU6375HLF22" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                End If
                
                   If Left(TmpChip, 10) = "AU6377ALF2" Then
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                      LBA = TmpLBA
                   End If
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
                
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
              
                
                 rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
               
                ClosePipe
                
                 If rv1 = 1 And rv2 <> 1 Then
                 Tester.Label9.Caption = "SMC FAIL"
                End If
                '============= SMC test begin =======================================
               
              
                 
                          '--- for SMC
                CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
                Call MsecDelay(0.1)
                      
                
                CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
                Call MsecDelay(0.1)
                ClosePipe
                
                
               
                  rv3 = CBWTest_New(2, rv2, VidName)
               
                Call LabelMenu(2, rv3, rv2)
                ClosePipe
                  If rv2 = 1 And rv3 <> 1 Then
                 Tester.Label9.Caption = "XD FAIL"
                End If
                
                
              
                
               '=============== SMC test END ==================================================
               
              
               rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
               
               If rv4 = 1 Then
                  rv4 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                  If rv4 <> 1 Then
                    rv4 = 2
                    Tester.Print "MS bus width Fail"
                    End If
               End If
                    
               
               
            
               ClosePipe
               Call LabelMenu(3, rv4, rv3)
               
               
               
 '======================== light test ======================
            Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            Tester.Print "LBA="; LBA
               
               
               
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
      If rv4 = 1 Then
          CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  for HID mode
          result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
          Call MsecDelay(0.4)
          
          CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
                            
    
          Call MsecDelay(1.5)
          CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
          If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
                  UsbSpeedTestResult = GPO_FAIL
                    rv4 = 2
                    Tester.Label9.Caption = "GPO FAIL " & LightOff
           End If
     End If
           Tester.Print rv4, " \\GPO,LED fail"
        If rv4 = 1 Then
        ' code begin
                     
                     Tester.Cls
                     Tester.Print "keypress test begin---------------"
                     Dim ReturnValue As Byte
                     
                     DeviceHandle = &HFFFF  'invalid handle initial value
                     Call MsecDelay(0.1)
                     ReturnValue = fnGetDeviceHandle(DeviceHandle)
                     Tester.Print ReturnValue; Space(5); ' 1: pass the other refer btnstatus.h
                     Tester.Print "DeviceHandle="; DevicehHandle
                     
                     If ReturnValue <> 1 Then
                           rv0 = UNKNOW       '---> HID mode unknow device mode
                           Call LabelMenu(0, rv0, 1)
                           Tester.Label9.Caption = "HID mode unknow device"
                          fnFreeDeviceHandle (DeviceHandle)
                         GoTo AU6377ALFResult
                     End If
                     
                     '=======================
                     '  key press test, it will return 8 when key up, GPI 6 must do low go hi action
                     '========================
                     
                
                     Do
                          CardResult = DO_WritePort(card, Channel_P1A, &H40) 'GPI6 : bit 6: pull high
                          Sleep (200)
                         CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' GPI6 : bit 6: pull low
                         Sleep (1000)
                       
                         ReturnValue = fnInquiryBtnStatus(DeviceHandle)
                         Tester.Print i; Space(5); "Key press value="; ReturnValue
                         i = i + 1
                     Loop While i < 3 And ReturnValue <> 10
                    ' fnFreeDeviceHandle (DeviceHandle)
                   '  fnFreeDeviceHandle (DeviceHandle)
                     
                     If ReturnValue <> 10 Then
                     
                      rv1 = 2
                     
                       Call LabelMenu(1, rv1, rv0)
                       Label9.Caption = "KeyPress Fail"
                       
                     End If
                              
                 
           End If
                
                
                Tester.Print rv0, " \\Can not Find HID deevice"
                Tester.Print rv1, " \\Key press Fail"
             '   Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
             '   Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
             '   Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
             '   Tester.Print "LBA="; LBA
                
AU6377ALFResult:
                        
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
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        
                        End If
                
               
                  
                End If
                
                
                
     '    CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off
     '    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
     '    Call MsecDelay(0.8)
     '    CardResult = DO_WritePort(card, Channel_P1A, &H40)
         
               
   End Sub

Public Sub AU6476CLF27TestSub()
'On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
Dim i As Long
Dim j As Long
              
'2010/9/3  copy from AU6476CLF24 purpose to solve HID mode unstable issue
'2012/2/22 Add XD R/W 128K 0x5A pattern

    TmpChip = ChipName
    '==================================== Switch assign ==========================================
                
                
    If PCI7248InitFinish = 0 Then
          PCI7248Exist
    End If
                
  
    CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  // force into reader mode
    'result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
    Call MsecDelay(0.2)
    'CardResult = DO_WritePort(card, Channel_P1A, &H0)
    
    'Call MsecDelay(0.3)
               
    CardResult = DO_WritePort(card, Channel_P1A, &H3F)
          
          
'======================== Begin test ============================================
    
    '//////////////////////////////////////////////////
    '
    '   no card insert
    '
    '/////////////////////////////////////////////////
                
    Dim VidName As String
    'Dim ChipString As String
    VidName = "vid_058f"
    'Call MsecDelay(0.3)
    WaitDevOn (VidName)
    Call MsecDelay(0.2)
    LBA = LBA + 1
              
    ClosePipe
    rv0 = CBWTest_New_no_card(0, 1, VidName)
    'Tester.print "a1"
    Call LabelMenu(0, rv0, 1)
    ClosePipe
    rv1 = CBWTest_New_no_card(1, rv0, VidName)
    'Tester.print "a2"
    Call LabelMenu(1, rv1, rv0)
    ClosePipe
               
    rv2 = CBWTest_New_no_card(2, rv1, VidName)
    'Tester.print "a3"
    Call LabelMenu(2, rv2, rv1)
    ClosePipe
    rv3 = CBWTest_New_no_card(3, rv2, VidName)
    'Tester.print "a4"
    ClosePipe
    Call LabelMenu(3, rv3, rv2)
                
 '================================= Test light off =============================
                
    Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 card change bit fail"
    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 card change bit fail"
    Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 card change bit fail"
    Tester.Print rv3, " \\MS :0 Unknow device, 1 pass ,2 card change bit fail"
                
    ' test chip
    ' CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
    If LightOff <> 255 Then
        UsbSpeedTestResult = GPO_FAIL
        rv0 = 2
    End If

    Tester.Print rv0, " \\LED Fail"
    
                
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
             
              
'====================================== Assing R/W test switch =====================================
                 
    If TestResult = "PASS" Then
        TestResult = ""
        
        CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
        Call MsecDelay(0.1)
        
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
        rv0 = CBWTest_New(0, 1, VidName)    ' sd slot
                 
        If rv0 = 1 Then
           rv0 = Read_SD_Speed(0, 0, 64, "4Bits")
           If rv0 <> 1 Then
                Tester.Print "SD bus width Fail"
                rv0 = 2
           End If
        End If
                
                   
        Call LabelMenu(0, rv0, 1)
        ClosePipe
        rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
        
        Call LabelMenu(1, rv1, rv0)
        ClosePipe
              
        rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
        
        If rv2 = 1 Then
            For i = 0 To 1
                rv2 = Write_Data_AU6377(LBA + i, 2, 65536)
                If rv2 <> 1 Then
                    Exit For
                End If
            Next
        End If
    
        If rv2 = 1 Then
            For i = 0 To 1
                rv2 = Read_Data(LBA + i, 2, 65536)
                
                If rv2 <> 1 Then
                    Exit For
                End If
                
                For j = 0 To 65535
                    If Pattern_AU6377(j) <> ReadData(j) Then
                        Tester.Print "LBA= " & LBA + i & " Cycle= " & j & " Value= " & Hex(ReadData(j))
                        rv2 = 3
                        Exit For
                    End If
                Next
            Next
        End If

        Call LabelMenu(2, rv2, rv1)
        ClosePipe
                
        If rv1 = 1 And rv2 <> 1 Then
            Tester.Label9.Caption = "SMC FAIL"
        End If
        
        '============= SMC test begin =======================================
               
              
                 
        '--- for SMC
        CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
        Call MsecDelay(0.1)
              
        
        CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
        Call MsecDelay(0.1)
        ClosePipe
               
        rv3 = CBWTest_New(2, rv2, VidName)
        Call LabelMenu(2, rv3, rv2)
        ClosePipe
        
        If rv2 = 1 And rv3 <> 1 Then
            Tester.Label9.Caption = "XD FAIL"
        End If
                
                
        '=============== SMC test END ==================================================
               
              
        rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
        
        If rv4 = 1 Then
        rv4 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
            If rv4 <> 1 Then
                rv4 = 2
                Tester.Print "MS bus width Fail"
            End If
        End If
        
        ClosePipe
        Call LabelMenu(3, rv4, rv3)
              
        '======================== light test ======================
        Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
        Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
        Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
        Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
        Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
        Tester.Print "LBA="; LBA
               
               
               
        '=================================================================================
        ' HID mode and reader mode ---> compositive device
        If rv4 = 1 Then
            CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  for HID mode
            'result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
            WaitDevOFF (VidName)
            
            CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
            
            WaitDevOn (VidName)
            Call MsecDelay(0.2)
            CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
            If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
                UsbSpeedTestResult = GPO_FAIL
                rv4 = 2
                Tester.Label9.Caption = "GPO FAIL " & LightOff
            End If
        End If
        
        Tester.Print rv4, " \\GPO,LED fail"
        
        If rv4 = 1 Then
        ' code begin
                     
            Tester.Cls
            Tester.Print "keypress test begin---------------"
            Dim ReturnValue As Byte
            
            DeviceHandle = &HFFFF  'invalid handle initial value
            Call MsecDelay(0.1)
            ReturnValue = fnGetDeviceHandle(DeviceHandle)
            Tester.Print ReturnValue; Space(5); ' 1: pass the other refer btnstatus.h
            Tester.Print "DeviceHandle="; DevicehHandle
            
            If ReturnValue <> 1 Then
                rv0 = UNKNOW       '---> HID mode unknow device mode
                Call LabelMenu(0, rv0, 1)
                Tester.Label9.Caption = "HID mode unknow device"
                fnFreeDeviceHandle (DeviceHandle)
                GoTo AU6377ALFResult
            End If
                     
            '=======================
            '  key press test, it will return 8 when key up, GPI 6 must do low go hi action
            '========================

            Do
                CardResult = DO_WritePort(card, Channel_P1A, &H40) 'GPI6 : bit 6: pull high
                Sleep (200)
                CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' GPI6 : bit 6: pull low
                Sleep (400)
                
                ReturnValue = fnInquiryBtnStatus(DeviceHandle)
                Tester.Print i; Space(5); "Key press value="; ReturnValue
                i = i + 1
            Loop While i < 3 And ReturnValue <> 10
            'fnFreeDeviceHandle (DeviceHandle)
            'fnFreeDeviceHandle (DeviceHandle)
                     
            If ReturnValue <> 10 Then
                rv1 = 2
                Call LabelMenu(1, rv1, rv0)
                Label9.Caption = "KeyPress Fail"
            End If
                              
                 
        End If
                
        Tester.Print rv0, " \\Can not Find HID deevice"
        Tester.Print rv1, " \\Key press Fail"
                
        
AU6377ALFResult:
        CardResult = DO_WritePort(card, Channel_P1A, &H80)
        WaitDevOFF (VidName)
        
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
        ElseIf rv4 = WRITE_FAIL Then
            MSWriteFail = MSWriteFail + 1
            TestResult = "MS_WF"
        ElseIf rv4 = READ_FAIL Then
            MSReadFail = MSReadFail + 1
            TestResult = "MS_RF"
        ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
            TestResult = "PASS"
        Else
            TestResult = "Bin2"
        End If
    End If
                
End Sub

Public Sub AU6476CLF06TestSub()
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
Dim HV_Flag As Boolean
Dim HV_Result As String
Dim LV_Result As String

               
'2010/9/3  copy from AU6476CLF24 purpose to solve HID mode unstable issue

HV_Flag = False
HV_Result = ""
LV_Result = ""
TmpChip = ChipName
'==================================== Switch assign ==========================================
ReaderExist = 0
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
Tester.Cls

  
ReTestLabel:
                  
                  
        If (HV_Flag = False) Then
            Call PowerSet2(1, "5.3", "0.5", 1, "5.3", "0.5", 1)
            Tester.Print "Begin HV Test ..."
        Else
            Call PowerSet2(1, "4.7", "0.5", 1, "4.7", "0.5", 1)
            Tester.Print vbCrLf & "Begin LV Test ..."
        End If
            
 
         CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  // force into reader mode
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.2)
         CardResult = DO_WritePort(card, Channel_P1A, &H0)
         
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H3F)
          
          
   
  
'======================== Begin test ============================================
                  
                'Call MsecDelay(1.6)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
               ' Dim ChipString As String
                VidName = "vid_058f"
                
                Call MsecDelay(0#)
                rv0 = WaitDevOn(VidName)
                Call MsecDelay(0.1)
                
                
                ClosePipe
                rv0 = CBWTest_New_no_card(0, 1, VidName)
                'Tester.print "a1"
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                rv1 = CBWTest_New_no_card(1, rv0, VidName)
               '  Tester.print "a2"
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
               
                rv2 = CBWTest_New_no_card(2, rv1, VidName)
               '  Tester.print "a3"
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
                
                rv3 = CBWTest_New_no_card(3, rv2, VidName)
                ' Tester.print "a4"
                ClosePipe
                Call LabelMenu(3, rv3, rv2)
                
 '================================= Test light off =============================
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv3, " \\MS :0 Unknow device, 1 pass ,2 card change bit fail"
                 
                
                
             
                
                ' test chip
                      '    CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If LightOff <> 255 Then
                            UsbSpeedTestResult = GPO_FAIL
                             rv0 = 2
                         End If
          
               Tester.Print rv0, " \\LED Fail"
                
                
                
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
               
                       
       
                 
              
'====================================== Assing R/W test switch =====================================
                 
                If TestResult = "PASS" Then
                   TestResult = ""
                   
                    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                   
              
                    
                   Call MsecDelay(0.1)
                 
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
                 rv0 = CBWTest_New(0, 1, VidName)    ' sd slot
                 
                 If rv0 = 1 Then
                    rv0 = Read_SD_Speed(0, 0, 64, "4Bits")
                    If rv0 <> 1 Then
                      Tester.Print "SD bus width Fail"
                      rv0 = 2
                    End If
                 End If
                   
                 
                 If rv0 = 1 And TmpChip = "AU6375HLF22" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                End If
                
                   If Left(TmpChip, 10) = "AU6377ALF2" Then
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                      LBA = TmpLBA
                   End If
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
                
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
              
                
                 rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
               
                ClosePipe
                
                 If rv1 = 1 And rv2 <> 1 Then
                 Tester.Label9.Caption = "SMC FAIL"
                End If
                '============= SMC test begin =======================================
               
              
                 
                          '--- for SMC
                CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
                Call MsecDelay(0.1)
                      
                
                CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
                Call MsecDelay(0.1)
                ClosePipe
                
                
               
                  rv3 = CBWTest_New(2, rv2, VidName)
               
                Call LabelMenu(2, rv3, rv2)
                ClosePipe
                  If rv2 = 1 And rv3 <> 1 Then
                 Tester.Label9.Caption = "XD FAIL"
                End If
                
                
              
                
               '=============== SMC test END ==================================================
               
              
               rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
               
               If rv4 = 1 Then
                  rv4 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                  If rv4 <> 1 Then
                    rv4 = 2
                    Tester.Print "MS bus width Fail"
                    End If
               End If
                    
               
               
            
               ClosePipe
               Call LabelMenu(3, rv4, rv3)
               
               
               
 '======================== light test ======================
            Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
            Tester.Print "LBA="; LBA
               
               
               
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
      If rv4 = 1 Then
          CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  for HID mode
          result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
          Call MsecDelay(0.4)
          
          CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
                            
    
          Call MsecDelay(1.5)
          CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
          If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
                  UsbSpeedTestResult = GPO_FAIL
                    rv4 = 2
                    Tester.Label9.Caption = "GPO FAIL " & LightOff
           End If
     End If
           Tester.Print rv4, " \\GPO,LED fail"
        If rv4 = 1 Then
        ' code begin
                     
                     Tester.Cls
                     Tester.Print "keypress test begin---------------"
                     Dim ReturnValue As Byte
                     
                     DeviceHandle = &HFFFF  'invalid handle initial value
                     Call MsecDelay(0.1)
                     ReturnValue = fnGetDeviceHandle(DeviceHandle)
                     Tester.Print ReturnValue; Space(5); ' 1: pass the other refer btnstatus.h
                     Tester.Print "DeviceHandle="; DevicehHandle
                     
                     If ReturnValue <> 1 Then
                           rv0 = UNKNOW       '---> HID mode unknow device mode
                           Call LabelMenu(0, rv0, 1)
                           Tester.Label9.Caption = "HID mode unknow device"
                          fnFreeDeviceHandle (DeviceHandle)
                         GoTo AU6377ALFResult
                     End If
                     
                     '=======================
                     '  key press test, it will return 8 when key up, GPI 6 must do low go hi action
                     '========================
                     
                
                     Do
                          CardResult = DO_WritePort(card, Channel_P1A, &H40) 'GPI6 : bit 6: pull high
                          Sleep (200)
                         CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' GPI6 : bit 6: pull low
                         Sleep (400)
                       
                         ReturnValue = fnInquiryBtnStatus(DeviceHandle)
                         Tester.Print i; Space(5); "Key press value="; ReturnValue
                         i = i + 1
                     Loop While i < 3 And ReturnValue <> 10
                    ' fnFreeDeviceHandle (DeviceHandle)
                   '  fnFreeDeviceHandle (DeviceHandle)
                     
                     If ReturnValue <> 10 Then
                     
                      rv1 = 2
                     
                       Call LabelMenu(1, rv1, rv0)
                       Label9.Caption = "KeyPress Fail"
                       
                     End If
                              
                 
           End If
                
                
                Tester.Print rv0, " \\Can not Find HID deevice"
                Tester.Print rv1, " \\Key press Fail"
             '   Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
             '   Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
             '   Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
             '   Tester.Print "LBA="; LBA
                
AU6377ALFResult:
                       
                Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
                CardResult = DO_WritePort(card, Channel_P1A, &H80)
                        
                If HV_Flag = False Then
                    If rv0 * rv1 * rv2 * rv3 * rv4 = 0 Then
                        HV_Result = "Bin2"
                        Tester.Print "HV Unknow"
                    ElseIf rv0 * rv1 * rv2 * rv3 * rv4 <> 1 Then
                        HV_Result = "Fail"
                        Tester.Print "HV Fail"
                    ElseIf rv0 * rv1 * rv2 * rv3 * rv4 = 1 Then
                        HV_Result = "PASS"
                        Tester.Print "HV PASS"
                    End If
                    rv0 = 0
                    rv1 = 0
                    rv2 = 0
                    rv3 = 0
                    rv4 = 0
                    HV_Flag = True
                    ReaderExist = 0
                    Call MsecDelay(0.2)
                    GoTo ReTestLabel
                Else
                    If rv0 * rv1 * rv2 * rv3 * rv4 = 0 Then
                        LV_Result = "Bin2"
                        Tester.Print "LV Unknow"
                    ElseIf rv0 * rv1 * rv2 * rv3 * rv4 <> 1 Then
                        LV_Result = "Fail"
                        Tester.Print "LV Fail"
                    ElseIf rv0 * rv1 * rv2 * rv3 * rv4 = 1 Then
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
                
      End If
   End Sub

Public Sub AU6476CLF07TestSub()
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
Dim HV_Flag As Boolean
Dim HV_Result As String
Dim LV_Result As String
Dim i As Long
Dim j As Long

               
'2010/9/3  copy from AU6476CLF24 purpose to solve HID mode unstable issue
'2012/2/21 Add XD R/W 128K 0x5A pattern for fiti RMA

    HV_Flag = False
    HV_Result = ""
    LV_Result = ""
    TmpChip = ChipName
    '==================================== Switch assign ==========================================
    ReaderExist = 0
                
    If PCI7248InitFinish = 0 Then
          PCI7248Exist
    End If
    Tester.Cls

  
ReTestLabel:
                  
                  
    If (HV_Flag = False) Then
        Call PowerSet2(1, "5.3", "0.5", 1, "5.3", "0.5", 1)
        Tester.Print "Begin HV Test ..."
    Else
        Call PowerSet2(1, "4.7", "0.5", 1, "4.7", "0.5", 1)
        Tester.Print vbCrLf & "Begin LV Test ..."
    End If
            
 
    CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  // force into reader mode
    'result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
    Call MsecDelay(0.2)
    'CardResult = DO_WritePort(card, Channel_P1A, &H0)
    'Call MsecDelay(0.3)
    CardResult = DO_WritePort(card, Channel_P1A, &H3F)

  
    '======================== Begin test ============================================
                  
    'Call MsecDelay(1.6)
    
    LBA = LBA + 1
    
    
    '//////////////////////////////////////////////////
    '
    '   no card insert
    '
    '/////////////////////////////////////////////////
    
    Dim VidName As String
    ' Dim ChipString As String
    VidName = "vid_058f"
    
    'Call MsecDelay(0.3)
    WaitDevOn (VidName)
    Call MsecDelay(0.2)
    
                
    ClosePipe
    rv0 = CBWTest_New_no_card(0, 1, VidName)
    'Tester.print "a1"
    Call LabelMenu(0, rv0, 1)
    ClosePipe
    rv1 = CBWTest_New_no_card(1, rv0, VidName)
    '  Tester.print "a2"
    Call LabelMenu(1, rv1, rv0)
    ClosePipe
    
    rv2 = CBWTest_New_no_card(2, rv1, VidName)
    '  Tester.print "a3"
    Call LabelMenu(2, rv2, rv1)
    ClosePipe
    
    rv3 = CBWTest_New_no_card(3, rv2, VidName)
    ' Tester.print "a4"
    ClosePipe
    Call LabelMenu(3, rv3, rv2)
                
    '================================= Test light off =============================
                
    Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 card change bit fail"
    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 card change bit fail"
    Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 card change bit fail"
    Tester.Print rv3, " \\MS :0 Unknow device, 1 pass ,2 card change bit fail"
                 
                
    'test chip
    'CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
    If LightOff <> 255 Then
        UsbSpeedTestResult = GPO_FAIL
        rv0 = 2
    End If
          
    Tester.Print rv0, " \\LED Fail"
               
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
            
              
    '====================================== Assing R/W test switch =====================================
                 
    If TestResult = "PASS" Then
        TestResult = ""
                   
        CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                
        Call MsecDelay(0.1)
                 
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
        rv0 = CBWTest_New(0, 1, VidName)    ' sd slot
                 
        If rv0 = 1 Then
           rv0 = Read_SD_Speed(0, 0, 64, "4Bits")
           If rv0 <> 1 Then
             Tester.Print "SD bus width Fail"
             rv0 = 2
           End If
        End If
                   
        Call LabelMenu(0, rv0, 1)
        ClosePipe
        rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
        
        Call LabelMenu(1, rv1, rv0)
        ClosePipe
                
        rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
        
        If rv2 = 1 Then
            For i = 0 To 1
                rv2 = Write_Data_AU6377(LBA + i, 2, 65536)
                If rv2 <> 1 Then
                    Exit For
                End If
            Next
        End If
    
        If rv2 = 1 Then
            For i = 0 To 1
                rv2 = Read_Data(LBA + i, 2, 65536)
                
                If rv2 <> 1 Then
                    Exit For
                End If
                
                For j = 0 To 65535
                    If Pattern_AU6377(j) <> ReadData(j) Then
                        Tester.Print "LBA= " & LBA + i & " Cycle= " & j & " Value= " & Hex(ReadData(j))
                        rv2 = 3
                        Exit For
                    End If
                Next
            Next
        End If
        
        Call LabelMenu(2, rv2, rv1)
        ClosePipe
                
        If rv1 = 1 And rv2 <> 1 Then
            Tester.Label9.Caption = "SMC FAIL"
        End If
                
        '============= SMC test begin =======================================
               
        '--- for SMC
        CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
        Call MsecDelay(0.1)
            
        CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
        Call MsecDelay(0.1)
        ClosePipe
        
        rv3 = CBWTest_New(2, rv2, VidName)
        
        Call LabelMenu(2, rv3, rv2)
        
        ClosePipe
        If rv2 = 1 And rv3 <> 1 Then
            Tester.Label9.Caption = "XD FAIL"
        End If
                
                
        '=============== SMC test END ==================================================
               
        rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
        
        If rv4 = 1 Then
           rv4 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
           If rv4 <> 1 Then
             rv4 = 2
             Tester.Print "MS bus width Fail"
             End If
        End If
            
        ClosePipe
        Call LabelMenu(3, rv4, rv3)

        '======================== light test ======================
        Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
        Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
        Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
        Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
        Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
        Tester.Print "LBA="; LBA

               
        '=================================================================================
        ' HID mode and reader mode ---> compositive device
        If rv4 = 1 Then
            CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  for HID mode
            WaitDevOFF (VidName)
            
            CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
            
            WaitDevOn (VidName)
            Call MsecDelay(0.2)
            CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
            If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
                UsbSpeedTestResult = GPO_FAIL
                rv4 = 2
                Tester.Label9.Caption = "GPO FAIL " & LightOff
            End If
        End If
        
        
        Tester.Print rv4, " \\GPO,LED fail"
        If rv4 = 1 Then
        ' code begin
                     
            Tester.Cls
            Tester.Print "keypress test begin---------------"
            Dim ReturnValue As Byte
            
            DeviceHandle = &HFFFF  'invalid handle initial value
            Call MsecDelay(0.1)
            ReturnValue = fnGetDeviceHandle(DeviceHandle)
            Tester.Print ReturnValue; Space(5); ' 1: pass the other refer btnstatus.h
            Tester.Print "DeviceHandle="; DevicehHandle
            
            If ReturnValue <> 1 Then
                rv0 = UNKNOW       '---> HID mode unknow device mode
                Call LabelMenu(0, rv0, 1)
                Tester.Label9.Caption = "HID mode unknow device"
                fnFreeDeviceHandle (DeviceHandle)
                GoTo AU6377ALFResult
            End If
                     
            '=======================
            '  key press test, it will return 8 when key up, GPI 6 must do low go hi action
            '========================
            
                
            Do
                CardResult = DO_WritePort(card, Channel_P1A, &H40) 'GPI6 : bit 6: pull high
                Sleep (200)
                CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' GPI6 : bit 6: pull low
                Sleep (400)
                
                ReturnValue = fnInquiryBtnStatus(DeviceHandle)
                Tester.Print i; Space(5); "Key press value="; ReturnValue
                i = i + 1
            Loop While i < 3 And ReturnValue <> 10
            'fnFreeDeviceHandle (DeviceHandle)
            'fnFreeDeviceHandle (DeviceHandle)
                         
            If ReturnValue <> 10 Then
                rv1 = 2
                Call LabelMenu(1, rv1, rv0)
                Label9.Caption = "KeyPress Fail"
            End If
                              
                 
        End If
                
                
        Tester.Print rv0, " \\Can not Find HID deevice"
        Tester.Print rv1, " \\Key press Fail"
                
AU6377ALFResult:
                       
        Call PowerSet2(1, "0.0", "0.5", 1, "0.0", "0.5", 1)
        CardResult = DO_WritePort(card, Channel_P1A, &H80)
        WaitDevOFF (VidName)
        
        If HV_Flag = False Then
            If rv0 * rv1 * rv2 * rv3 * rv4 = 0 Then
                HV_Result = "Bin2"
                Tester.Print "HV Unknow"
            ElseIf rv0 * rv1 * rv2 * rv3 * rv4 <> 1 Then
                HV_Result = "Fail"
                Tester.Print "HV Fail"
            ElseIf rv0 * rv1 * rv2 * rv3 * rv4 = 1 Then
                HV_Result = "PASS"
                Tester.Print "HV PASS"
            End If
            rv0 = 0
            rv1 = 0
            rv2 = 0
            rv3 = 0
            rv4 = 0
            HV_Flag = True
            ReaderExist = 0
            GoTo ReTestLabel
        Else
            If rv0 * rv1 * rv2 * rv3 * rv4 = 0 Then
                LV_Result = "Bin2"
                Tester.Print "LV Unknow"
            ElseIf rv0 * rv1 * rv2 * rv3 * rv4 <> 1 Then
                LV_Result = "Fail"
                Tester.Print "LV Fail"
            ElseIf rv0 * rv1 * rv2 * rv3 * rv4 = 1 Then
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
                
    End If
      
End Sub
   
Public Sub AU6476CLTestSub()
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
               
   
TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
  
                  
 If Left(TmpChip, 8) = "AU6378AL" Or Left(TmpChip, 8) = "AU6476CL" Then
 
         CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  // force into reader mode
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.2)
         CardResult = DO_WritePort(card, Channel_P1A, &H0)
         
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H3F)
          
          
  End If
  
  
  '========================== AU6377 new board switch assign ment  ============
If TmpChip = "AU6377ALF21" Then ' this for new board  and internalrom
         RomSelector = &H10  '-------- this is for MS in pin
  End If
  
  
  If TmpChip = "AU6377ALF24" Then ' this for new board  and internalrom
         RomSelector = &H10
  End If
  
  If TmpChip = "AU6377ALF25" Then ' this for new board  and internalrom
         RomSelector = &H0
  End If
         
         
  If Left(TmpChip, 10) = "AU6377ALF2" Then
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector)  ' 1111 1110
         CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' 1111 1110
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector) ' 5th bit is rom selector, High is internal rom
  End If
  
  


  
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1.3)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
                Tester.Print LBA
                If TmpChip = "AU6377ALF25" Or TmpChip = "AU6378ALF21" Then
                  VidName = "vid_1984"
                Else
                 VidName = "vid_058f"
                End If
                
              
                ClosePipe
                 rv0 = CBWTest_New_no_card(0, 1, VidName)
                'Tester.print "a1"
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                rv1 = CBWTest_New_no_card(1, rv0, VidName)
               '  Tester.print "a2"
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
               
               rv2 = CBWTest_New_no_card(2, rv1, VidName)
               '  Tester.print "a3"
                 Call LabelMenu(2, rv2, rv1)
                ClosePipe
                  rv3 = CBWTest_New_no_card(3, rv2, VidName)
                ' Tester.print "a4"
                ClosePipe
               Call LabelMenu(3, rv3, rv2)
                
 '================================= Test light off =============================
                
              
                
                
                If TmpChip = "AU6476CLF20" Then
                
                ' test chip
                      '    CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If LightOff <> 255 Then
                            UsbSpeedTestResult = GPO_FAIL
                             rv0 = 2
                         End If
          
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
               
                       
       
                 
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv3, " \\MS :0 Unknow device, 1 pass ,2 card change bit fail"
                 
'====================================== Assing R/W test switch =====================================
                 
                If TestResult = "PASS" Then
                   TestResult = ""
                   If ChipName = "AU6476CLF20" Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                   End If
              
                    
                   Call MsecDelay(0.1)
                 
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
                 rv0 = CBWTest_New(0, 1, VidName)     ' sd slot
                 
                 If rv0 = 1 And TmpChip = "AU6375HLF22" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                End If
                
                   If Left(TmpChip, 10) = "AU6377ALF2" Then
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                      LBA = TmpLBA
                   End If
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
                
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
              
                
                 rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
               
                ClosePipe
                
                 If rv1 = 1 And rv2 <> 1 Then
                 Tester.Label9.Caption = "SMC FAIL"
                End If
                '============= SMC test begin =======================================
               
              
                 
                          '--- for SMC
                CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
                Call MsecDelay(0.1)
                      
                
                CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
                Call MsecDelay(0.1)
                ClosePipe
                
                
               
                  rv3 = CBWTest_New(2, rv2, VidName)
               
                Call LabelMenu(2, rv3, rv2)
                ClosePipe
                  If rv2 = 1 And rv3 <> 1 Then
                 Tester.Label9.Caption = "XD FAIL"
                End If
                
                
              
                
               '=============== SMC test END ==================================================
               
              
               rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
            
               ClosePipe
               Call LabelMenu(3, rv4, rv3)
               
 '======================== light test ======================
               
               
               
               
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
      If rv4 = 1 Then
          CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off  for HID mode
          result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
          Call MsecDelay(0.4)
          
          CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
                            
    
          Call MsecDelay(1.2)
          CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
          If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
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
                         GoTo AU6377ALFResult
                     End If
                     
                     '=======================
                     '  key press test, it will return 8 when key up, GPI 6 must do low go hi action
                     '========================
                     
                
                     Do
                          CardResult = DO_WritePort(card, Channel_P1A, &H40) 'GPI6 : bit 6: pull high
                          Sleep (200)
                         CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' GPI6 : bit 6: pull low
                         Sleep (1000)
                       
                         ReturnValue = fnInquiryBtnStatus(DeviceHandle)
                         Tester.Print i; Space(5); "Key press value="; ReturnValue
                         i = i + 1
                     Loop While i < 3 And ReturnValue <> 11
                    ' fnFreeDeviceHandle (DeviceHandle)
                   '  fnFreeDeviceHandle (DeviceHandle)
                     
                     If ReturnValue <> 11 Then
                     
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
                
AU6377ALFResult:
                        
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
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        
                        End If
                
               
                  
                End If
                
                
                
     '    CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off
     '    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
     '    Call MsecDelay(0.8)
     '    CardResult = DO_WritePort(card, Channel_P1A, &H40)
         
                
                
               
   End Sub

Public Sub AU6476FLTestSub()
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
               
   
TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
  
                  
 If Left(TmpChip, 8) = "AU6378AL" Or Left(TmpChip, 8) = "AU6476FL" Then
 
         CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  // force into reader mode
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.2)
         CardResult = DO_WritePort(card, Channel_P1A, &H0)
         
         Call MsecDelay(0.3)
                    
        
          
          
  End If
  
  
 
  
  
  


  
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1.3)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
                Tester.Print LBA
                If TmpChip = "AU6377ALF25" Or TmpChip = "AU6378ALF21" Then
                  VidName = "vid_1984"
                Else
                 VidName = "vid_058f"
                End If
                
              
          
                   TestResult = ""
                   If ChipName = "AU6476FLF20" Then
             '       CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                   End If
              
                    
                   Call MsecDelay(0.1)
                 
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
                 rv0 = CBWTest_New(0, 1, VidName)     ' sd slot
                 
                 If rv0 = 1 And TmpChip = "AU6375HLF22" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                End If
                
                   If Left(TmpChip, 10) = "AU6377ALF2" Then
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                      LBA = TmpLBA
                   End If
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
                
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
              
                
                 rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
               
                ClosePipe
                
                 If rv1 = 1 And rv2 <> 1 Then
                 Tester.Label9.Caption = "SMC FAIL"
                End If
                '============= SMC test begin =======================================
               
              
                 
                          '--- for SMC
                CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
                Call MsecDelay(0.1)
                      
                
                CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
                Call MsecDelay(0.1)
                ClosePipe
                
                
               
                  rv3 = CBWTest_New(2, rv2, VidName)
               
                Call LabelMenu(2, rv3, rv2)
                ClosePipe
                  If rv2 = 1 And rv3 <> 1 Then
                 Tester.Label9.Caption = "XD FAIL"
                End If
                
                
              
                
               '=============== SMC test END ==================================================
               
              
               rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
            
               ClosePipe
               Call LabelMenu(3, rv4, rv3)
               
 '======================== light test ======================
               
               
               
               
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
      If rv4 = 1 Then
          
          
          CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
                            
    
          Call MsecDelay(1.2)
          CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
          If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
                  UsbSpeedTestResult = GPO_FAIL
                    rv4 = 2
                    Tester.Label9.Caption = "GPO FAIL " & LightOff
           End If
     End If
          
       
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print "LBA="; LBA
                
AU6377ALFResult:
                        
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
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        
                        End If
                
               
                  
                
                
                
                
     '    CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off
     '    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
     '    Call MsecDelay(0.8)
     '    CardResult = DO_WritePort(card, Channel_P1A, &H40)
         
                
                
               
   End Sub
   
Public Sub AU6476FL24TestSub()
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
               
   
TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
  
                  
 If Left(TmpChip, 8) = "AU6378AL" Or Left(TmpChip, 8) = "AU6476FL" Then
 
         CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  // force into reader mode
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.2)
         CardResult = DO_WritePort(card, Channel_P1A, &H0)
         
         Call MsecDelay(0.3)
                    
        
          
          
  End If
  
  
 
  
  
  


  
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1.3)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
                Tester.Print LBA
                If TmpChip = "AU6377ALF25" Or TmpChip = "AU6378ALF21" Then
                  VidName = "vid_1984"
                Else
                 VidName = "vid_058f"
                End If
                
              
          
                   TestResult = ""
                   If ChipName = "AU6476FLF20" Then
             '       CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                   End If
              
                    
                   Call MsecDelay(0.1)
                 
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
                 rv0 = CBWTest_New(0, 1, VidName)    ' sd slot
                 
                 If rv0 = 1 Then
                 rv0 = Read_SD_Speed(0, 0, 64, "4Bits")
                    If rv0 <> 1 Then
                        rv0 = 2
                        Tester.Print "SD bus width Fail"
                    End If
                 End If
                     
                 
                 
                 If rv0 = 1 And TmpChip = "AU6375HLF22" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                End If
                
                   If Left(TmpChip, 10) = "AU6377ALF2" Then
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                      LBA = TmpLBA
                   End If
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
                
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
              
                
                 rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
               
                ClosePipe
                
                 If rv1 = 1 And rv2 <> 1 Then
                 Tester.Label9.Caption = "SMC FAIL"
                End If
                '============= SMC test begin =======================================
               
              
                 
                          '--- for SMC
                CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
                Call MsecDelay(0.1)
                      
                
                CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
                Call MsecDelay(0.1)
                ClosePipe
                
                
               
                  rv3 = CBWTest_New(2, rv2, VidName)
               
                Call LabelMenu(2, rv3, rv2)
                ClosePipe
                  If rv2 = 1 And rv3 <> 1 Then
                 Tester.Label9.Caption = "XD FAIL"
                End If
                
                
              
                
               '=============== SMC test END ==================================================
               
              
               rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
               
               If rv4 = 1 Then
               rv4 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
               If rv4 <> 1 Then
                  rv4 = 2
                  Tester.Print "MS Bus width Fail"
               End If
               End If
            
               
            
               ClosePipe
               Call LabelMenu(3, rv4, rv3)
               
 '======================== light test ======================
               
               
               
               
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
      If rv4 = 1 Then
          
          
          CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
                            
    
          Call MsecDelay(1.2)
          CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
          If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
                  UsbSpeedTestResult = GPO_FAIL
                    rv4 = 2
                    Tester.Label9.Caption = "GPO FAIL " & LightOff
           End If
     End If
          
       
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print "LBA="; LBA
                
AU6377ALFResult:
                        
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
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        
                        End If
                
               
                  
                
                
                
                
     '    CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off
     '    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
     '    Call MsecDelay(0.8)
     '    CardResult = DO_WritePort(card, Channel_P1A, &H40)
         
                
                
               
   End Sub
   
Public Sub AU6476FL21TestSub()
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
               
   
TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
  
                  
 If Left(TmpChip, 8) = "AU6378AL" Or Left(TmpChip, 8) = "AU6476FL" Then
 
         CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  // force into reader mode
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.2)
         CardResult = DO_WritePort(card, Channel_P1A, &H0)
         
         Call MsecDelay(0.3)
                    
        
          
          
  End If
  
  
 
  
  
  


  
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1.3)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
                Tester.Print LBA
                If TmpChip = "AU6377ALF25" Or TmpChip = "AU6378ALF21" Then
                  VidName = "vid_1984"
                Else
                 VidName = "vid_058f"
                End If
                
              
          
                   TestResult = ""
                   If ChipName = "AU6476FLF20" Then
             '       CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                   End If
              
                    
                   Call MsecDelay(0.1)
                 
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
                 rv0 = CBWTest_New_SD_Speed(0, 1, VidName, "4Bits")    ' sd slot
                 
                 If rv0 = 1 And TmpChip = "AU6375HLF22" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                End If
                
                   If Left(TmpChip, 10) = "AU6377ALF2" Then
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                      LBA = TmpLBA
                   End If
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
                
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
              
                
                 rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
               
                ClosePipe
                
                 If rv1 = 1 And rv2 <> 1 Then
                 Tester.Label9.Caption = "SMC FAIL"
                End If
                '============= SMC test begin =======================================
               
              
                 
                          '--- for SMC
                CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
                Call MsecDelay(0.1)
                      
                
                CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
                Call MsecDelay(0.1)
                ClosePipe
                
                
               
                  rv3 = CBWTest_New(2, rv2, VidName)
               
                Call LabelMenu(2, rv3, rv2)
                ClosePipe
                  If rv2 = 1 And rv3 <> 1 Then
                 Tester.Label9.Caption = "XD FAIL"
                End If
                
                
              
                
               '=============== SMC test END ==================================================
               
              
               rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
            
               ClosePipe
               Call LabelMenu(3, rv4, rv3)
               
 '======================== light test ======================
               
               
               
               
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
      If rv4 = 1 Then
          
          
          CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
                            
    
          Call MsecDelay(1.2)
          CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
          If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
                  UsbSpeedTestResult = GPO_FAIL
                    rv4 = 2
                    Tester.Label9.Caption = "GPO FAIL " & LightOff
           End If
     End If
          
       
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print "LBA="; LBA
                
AU6377ALFResult:
                        
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
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        
                        End If
                
               
                  
                
                
                
                
     '    CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off
     '    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
     '    Call MsecDelay(0.8)
     '    CardResult = DO_WritePort(card, Channel_P1A, &H40)
         
                
                
               
   End Sub
Public Sub AU6476DLTestSub()
If ChipName = "AU6476DLF20" Then
    Call AU6476DLF20TestSub
End If

If ChipName = "AU6476DLF21" Then
    Call AU6476DLF21TestSub
End If

If ChipName = "AU6476DLF22" Then
    Call AU6476DLF22TestSub
End If

If ChipName = "AU6476DLF24" Then
    Call AU6476DLF24TestSub
End If


End Sub

Public Sub AU6476DLF20TestSub()
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
               
   
TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
  
                  
 If Left(TmpChip, 8) = "AU6378AL" Or Left(TmpChip, 8) = "AU6476DL" Then
 
         CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  // force into reader mode
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.2)
         CardResult = DO_WritePort(card, Channel_P1A, &H0)
         
         Call MsecDelay(0.3)
                    
        ' CardResult = DO_WritePort(card, Channel_P1A, &H3F)
          
          
  End If
  
  
  '========================== AU6377 new board switch assign ment  ============
If TmpChip = "AU6377ALF21" Then ' this for new board  and internalrom
         RomSelector = &H10  '-------- this is for MS in pin
  End If
  
  
  If TmpChip = "AU6377ALF24" Then ' this for new board  and internalrom
         RomSelector = &H10
  End If
  
  If TmpChip = "AU6377ALF25" Then ' this for new board  and internalrom
         RomSelector = &H0
  End If
         
         
  If Left(TmpChip, 10) = "AU6377ALF2" Then
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector)  ' 1111 1110
         CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' 1111 1110
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector) ' 5th bit is rom selector, High is internal rom
  End If
  
  


  
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1.3)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
                Tester.Print LBA
               
                 VidName = "vid_18e3"
                 
                
               
                
                
                If TmpChip = "AU6476DLF20" Then
                
                ' test chip
                      '    CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If LightOff <> 254 And LightOff <> 255 Then
                            UsbSpeedTestResult = GPO_FAIL
                             rv0 = 2
                         End If
          
                End If
                
                
             
'====================================== Assing R/W test switch =====================================
                 
                 
                   TestResult = ""
                   If ChipName = "AU6476DLF20" Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                   End If
              
                    
                   Call MsecDelay(0.1)
                 
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
                 rv0 = CBWTest_New(0, 1, VidName)     ' sd slot
                 
                 If rv0 = 1 And TmpChip = "AU6375HLF22" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                End If
                
                   If Left(TmpChip, 10) = "AU6377ALF2" Then
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                      LBA = TmpLBA
                   End If
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
                
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
              
                
                 rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
               
                ClosePipe
                
                 If rv1 = 1 And rv2 <> 1 Then
                 Tester.Label9.Caption = "SMC FAIL"
                End If
                '============= SMC test begin =======================================
               
              
                 
                          '--- for SMC
                CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
                Call MsecDelay(0.1)
                      
                
                CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
                Call MsecDelay(0.1)
                ClosePipe
                
                
               
                  rv3 = CBWTest_New(2, rv2, VidName)
               
                Call LabelMenu(2, rv3, rv2)
                ClosePipe
                  If rv2 = 1 And rv3 <> 1 Then
                 Tester.Label9.Caption = "XD FAIL"
                End If
                
                
              
                
               '=============== SMC test END ==================================================
               
              
               rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
            
               ClosePipe
               Call LabelMenu(3, rv4, rv3)
               
 '======================== light test ======================
               
               
               
               
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
      If rv4 = 1 Then
          CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 high, 7th bit of control 1, and pwr off  for HID mode
          result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
          Call MsecDelay(0.4)
          
          CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
                            
    
          Call MsecDelay(1.2)
          CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
          If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
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
                         GoTo AU6377ALFResult
                     End If
                     
                     '=======================
                     '  key press test, it will return 8 when key up, GPI 6 must do low go hi action
                     '========================
                     
                
                     Do
                          CardResult = DO_WritePort(card, Channel_P1A, &H40) 'GPI6 : bit 6: pull high
                          Sleep (200)
                         CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' GPI6 : bit 6: pull low
                         Sleep (1000)
                       
                         ReturnValue = fnInquiryBtnStatus(DeviceHandle)
                         Tester.Print i; Space(5); "Key press value="; ReturnValue
                         i = i + 1
                     Loop While i < 3 And ReturnValue <> 11
                    ' fnFreeDeviceHandle (DeviceHandle)
                   '  fnFreeDeviceHandle (DeviceHandle)
                     
                     If ReturnValue <> 11 Then
                     
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
                
AU6377ALFResult:
                        
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
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        
                        End If
                
               
                  
               
                
                
                
     '    CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off
     '    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
     '    Call MsecDelay(0.8)
     '    CardResult = DO_WritePort(card, Channel_P1A, &H40)
         
                
                
               
   End Sub
   
Public Sub AU6476DLF21TestSub()
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
               
   
TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
  
                  
 If Left(TmpChip, 8) = "AU6378AL" Or Left(TmpChip, 8) = "AU6476DL" Then
 
         CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  // force into reader mode
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.2)
         CardResult = DO_WritePort(card, Channel_P1A, &H0)
         
         Call MsecDelay(0.3)
                    
        ' CardResult = DO_WritePort(card, Channel_P1A, &H3F)
          
          
  End If
  
  
  '========================== AU6377 new board switch assign ment  ============
If TmpChip = "AU6377ALF21" Then ' this for new board  and internalrom
         RomSelector = &H10  '-------- this is for MS in pin
  End If
  
  
  If TmpChip = "AU6377ALF24" Then ' this for new board  and internalrom
         RomSelector = &H10
  End If
  
  If TmpChip = "AU6377ALF25" Then ' this for new board  and internalrom
         RomSelector = &H0
  End If
         
         
  If Left(TmpChip, 10) = "AU6377ALF2" Then
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector)  ' 1111 1110
         CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' 1111 1110
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector) ' 5th bit is rom selector, High is internal rom
  End If
  
  


  
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1.3)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
                Tester.Print LBA
               
                 VidName = "vid_18e3"
                 
                
               
                
                
                If TmpChip = "AU6476DLF20" Then
                
                ' test chip
                      '    CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If LightOff <> 254 And LightOff <> 255 Then
                            UsbSpeedTestResult = GPO_FAIL
                             rv0 = 2
                         End If
          
                End If
                
                
             
'====================================== Assing R/W test switch =====================================
                 
                 
                   TestResult = ""
                   If ChipName = "AU6476DLF20" Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                   End If
              
                    
                   Call MsecDelay(0.1)
                 
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
                 rv0 = CBWTest_New(0, 1, VidName)     ' sd slot
                 
                 If rv0 = 1 And TmpChip = "AU6375HLF22" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                End If
                
                   If Left(TmpChip, 10) = "AU6377ALF2" Then
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                      LBA = TmpLBA
                   End If
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
                
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
              
                
                 rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
               
                ClosePipe
                
                 If rv1 = 1 And rv2 <> 1 Then
                 Tester.Label9.Caption = "SMC FAIL"
                End If
                '============= SMC test begin =======================================
               
              
                 
                          '--- for SMC
                CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
                Call MsecDelay(0.1)
                      
                
                CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
                Call MsecDelay(0.1)
                ClosePipe
                
                
               
                  rv3 = CBWTest_New(2, rv2, VidName)
               
                Call LabelMenu(2, rv3, rv2)
                ClosePipe
                  If rv2 = 1 And rv3 <> 1 Then
                 Tester.Label9.Caption = "XD FAIL"
                End If
                
                
              
                
               '=============== SMC test END ==================================================
               
              
               rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
            
               ClosePipe
               Call LabelMenu(3, rv4, rv3)
               
 '======================== light test ======================
               
               
               
               
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
      If rv4 = 1 Then
          CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 high, 7th bit of control 1, and pwr off  for HID mode
          result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
          Call MsecDelay(0.4)
          
          CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
                            
    
          Call MsecDelay(1.2)
          CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
          If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
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
                         GoTo AU6377ALFResult
                     End If
                     
                     '=======================
                     '  key press test, it will return 8 when key up, GPI 6 must do low go hi action
                     '========================
                     
                
                     Do
                          CardResult = DO_WritePort(card, Channel_P1A, &H40) 'GPI6 : bit 6: pull high
                          Sleep (200)
                         CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' GPI6 : bit 6: pull low
                         Sleep (1000)
                       
                         ReturnValue = fnInquiryBtnStatus(DeviceHandle)
                         Tester.Print i; Space(5); "Key press value="; ReturnValue
                         i = i + 1
                     Loop While i < 3 And ReturnValue <> 10
                    ' fnFreeDeviceHandle (DeviceHandle)
                   '  fnFreeDeviceHandle (DeviceHandle)
                     
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
                
AU6377ALFResult:
                        
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
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        
                        End If
                
               
                  
               
                
                
                
     '    CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off
     '    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
     '    Call MsecDelay(0.8)
     '    CardResult = DO_WritePort(card, Channel_P1A, &H40)
         
                
                
               
   End Sub
Public Sub AU6476DLF22TestSub()
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
               
   
TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
  
                  
 If Left(TmpChip, 8) = "AU6378AL" Or Left(TmpChip, 8) = "AU6476DL" Then
 
         CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  // force into reader mode
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.2)
         CardResult = DO_WritePort(card, Channel_P1A, &H0)
         
         Call MsecDelay(0.3)
                    
        ' CardResult = DO_WritePort(card, Channel_P1A, &H3F)
          
          
  End If
  
  
  '========================== AU6377 new board switch assign ment  ============
If TmpChip = "AU6377ALF21" Then ' this for new board  and internalrom
         RomSelector = &H10  '-------- this is for MS in pin
  End If
  
  
  If TmpChip = "AU6377ALF24" Then ' this for new board  and internalrom
         RomSelector = &H10
  End If
  
  If TmpChip = "AU6377ALF25" Then ' this for new board  and internalrom
         RomSelector = &H0
  End If
         
         
  If Left(TmpChip, 10) = "AU6377ALF2" Then
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector)  ' 1111 1110
         CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' 1111 1110
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector) ' 5th bit is rom selector, High is internal rom
  End If
  
  


  
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1.3)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
                Tester.Print LBA
               
                 VidName = "vid_18e3"
                 
                
               
                
                
                If TmpChip = "AU6476DLF20" Then
                
                ' test chip
                      '    CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If LightOff <> 254 And LightOff <> 255 Then
                            UsbSpeedTestResult = GPO_FAIL
                             rv0 = 2
                         End If
          
                End If
                
                
             
'====================================== Assing R/W test switch =====================================
                 
                 
                   TestResult = ""
                   If ChipName = "AU6476DLF20" Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                   End If
              
                    
                   Call MsecDelay(0.1)
                 
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
                 rv0 = CBWTest_New_SD_Speed(0, 1, VidName, "8Bits")    ' sd slot
                 
                 If rv0 = 1 And TmpChip = "AU6375HLF22" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                End If
                
                   If Left(TmpChip, 10) = "AU6377ALF2" Then
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                      LBA = TmpLBA
                   End If
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
                
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
              
                
                 rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
               
                ClosePipe
                
                 If rv1 = 1 And rv2 <> 1 Then
                 Tester.Label9.Caption = "SMC FAIL"
                End If
                '============= SMC test begin =======================================
               
              
                 
                          '--- for SMC
                CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
                Call MsecDelay(0.1)
                      
                
                CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
                Call MsecDelay(0.1)
                ClosePipe
                
                
               
                  rv3 = CBWTest_New(2, rv2, VidName)
               
                Call LabelMenu(2, rv3, rv2)
                ClosePipe
                  If rv2 = 1 And rv3 <> 1 Then
                 Tester.Label9.Caption = "XD FAIL"
                End If
                
                
              
                
               '=============== SMC test END ==================================================
               
              
               rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
            
               ClosePipe
               Call LabelMenu(3, rv4, rv3)
               
 '======================== light test ======================
               
               
               
               
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
      If rv4 = 1 Then
          CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 high, 7th bit of control 1, and pwr off  for HID mode
          result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
          Call MsecDelay(0.4)
          
          CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
                            
    
          Call MsecDelay(1.2)
          CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
          If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
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
                         GoTo AU6377ALFResult
                     End If
                     
                     '=======================
                     '  key press test, it will return 8 when key up, GPI 6 must do low go hi action
                     '========================
                     
                
                     Do
                          CardResult = DO_WritePort(card, Channel_P1A, &H40) 'GPI6 : bit 6: pull high
                          Sleep (200)
                         CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' GPI6 : bit 6: pull low
                         Sleep (1000)
                       
                         ReturnValue = fnInquiryBtnStatus(DeviceHandle)
                         Tester.Print i; Space(5); "Key press value="; ReturnValue
                         i = i + 1
                     Loop While i < 3 And ReturnValue <> 10
                    ' fnFreeDeviceHandle (DeviceHandle)
                   '  fnFreeDeviceHandle (DeviceHandle)
                     
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
                
AU6377ALFResult:
                        
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
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        
                        End If
                
               
                  
               
                
                
                
     '    CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off
     '    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
     '    Call MsecDelay(0.8)
     '    CardResult = DO_WritePort(card, Channel_P1A, &H40)
         
                
                
               
   End Sub
Public Sub AU6476YLF26TestSub()
'On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
Dim i As Long
Dim j As Long
   
    TmpChip = ChipName
    '==================================== Switch assign ==========================================
                
                
    If PCI7248InitFinish = 0 Then
          PCI7248Exist
    End If
                
    CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  // force into reader mode
    Call MsecDelay(0.2)
    CardResult = DO_WritePort(card, Channel_P1A, &H0)
    
    'Call MsecDelay(0.3)
    
    '======================== Begin test ============================================
    '//////////////////////////////////////////////////
    '
    '   no card insert
    '
    '/////////////////////////////////////////////////
    LBA = LBA + 1
    
    Dim VidName As String
    Tester.Print LBA
    VidName = "vid_18e3"
    'Call MsecDelay(0.3)
    WaitDevOn (VidName)
    Call MsecDelay(0.2)
             
    '====================================== Assing R/W test switch =====================================
                 
    TestResult = ""
    
    Call MsecDelay(0.1)
                 
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
    rv0 = CBWTest_New(0, 1, VidName)    ' sd slot
    
    If rv0 = 1 Then
        rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
       
        If rv0 <> 1 Then
            rv0 = 2
            Tester.Print "SD bus width Fail"
        End If
    End If
                
    Call LabelMenu(0, rv0, 1)
    ClosePipe
    rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
    
    Call LabelMenu(1, rv1, rv0)
    ClosePipe
              
    rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
    
    If rv2 = 1 Then
        For i = 0 To 1
            rv2 = Write_Data_AU6377(LBA + i, 2, 65536)
            If rv2 <> 1 Then
                Exit For
            End If
        Next
    End If

    If rv2 = 1 Then
        For i = 0 To 1
            rv2 = Read_Data(LBA + i, 2, 65536)
            
            If rv2 <> 1 Then
                Exit For
            End If
            
            For j = 0 To 65535
                If Pattern_AU6377(j) <> ReadData(j) Then
                    Tester.Print "LBA= " & LBA + i & " Cycle= " & j & " Value= " & Hex(ReadData(j))
                    rv2 = 3
                    Exit For
                End If
            Next
        Next
    End If
    
    Call LabelMenu(2, rv2, rv1)
    
    ClosePipe
    
    If rv1 = 1 And rv2 <> 1 Then
        Tester.Label9.Caption = "SMC FAIL"
    End If
    '============= SMC test begin =======================================
    '--- for SMC
    CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
    Call MsecDelay(0.1)
          
    CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
    Call MsecDelay(0.1)
    ClosePipe
               
    rv3 = CBWTest_New(2, rv2, VidName)
    
    Call LabelMenu(2, rv3, rv2)
    ClosePipe
    If rv2 = 1 And rv3 <> 1 Then
        Tester.Label9.Caption = "XD FAIL"
    End If
                 
    '=============== SMC test END ==================================================
    rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
            
    If rv4 = 1 Then
        rv4 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
        If rv4 <> 1 Then
            rv4 = 2
            Tester.Print "MS bus width Fail"
        End If
    End If
         
    ClosePipe
    Call LabelMenu(3, rv4, rv3)
               
               
    '=================================================================================
    ' HID mode and reader mode ---> compositive device
    If rv4 = 1 Then
        CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 high, 7th bit of control 1, and pwr off  for HID mode
        WaitDevOFF (VidName)
        
        CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
        
        WaitDevOn (VidName)
        Call MsecDelay(0.2)
        CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
        
        If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
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
            GoTo AU6377ALFResult
        End If
                     
        '=======================
        '  key press test, it will return 8 when key up, GPI 6 must do low go hi action
        '========================
        Do
            CardResult = DO_WritePort(card, Channel_P1A, &H40) 'GPI6 : bit 6: pull high
            Sleep (200)
            CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' GPI6 : bit 6: pull low
            Sleep (400)
            
            ReturnValue = fnInquiryBtnStatus(DeviceHandle)
            Tester.Print i; Space(5); "Key press value="; ReturnValue
            i = i + 1
        Loop While i < 3 And ReturnValue <> 10
        'fnFreeDeviceHandle (DeviceHandle)
        'fnFreeDeviceHandle (DeviceHandle)
                     
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
                
AU6377ALFResult:
    CardResult = DO_WritePort(card, Channel_P1A, &H80)
    WaitDevOFF (VidName)
    
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
     ElseIf rv4 = WRITE_FAIL Then
        MSWriteFail = MSWriteFail + 1
        TestResult = "MS_WF"
    ElseIf rv4 = READ_FAIL Then
        MSReadFail = MSReadFail + 1
        TestResult = "MS_RF"
    ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
         TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
    
End Sub
   
Public Sub AU6476DLF24TestSub()
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
               
   
TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
  
                  
 If Left(TmpChip, 8) = "AU6378AL" Or Left(TmpChip, 8) = "AU6476DL" Then
 
         CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  // force into reader mode
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.2)
         CardResult = DO_WritePort(card, Channel_P1A, &H0)
         
         Call MsecDelay(0.3)
                    
        ' CardResult = DO_WritePort(card, Channel_P1A, &H3F)
          
          
  End If
  
  
  '========================== AU6377 new board switch assign ment  ============
If TmpChip = "AU6377ALF21" Then ' this for new board  and internalrom
         RomSelector = &H10  '-------- this is for MS in pin
  End If
  
  
  If TmpChip = "AU6377ALF24" Then ' this for new board  and internalrom
         RomSelector = &H10
  End If
  
  If TmpChip = "AU6377ALF25" Then ' this for new board  and internalrom
         RomSelector = &H0
  End If
         
         
  If Left(TmpChip, 10) = "AU6377ALF2" Then
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector)  ' 1111 1110
         CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' 1111 1110
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H6F + RomSelector) ' 5th bit is rom selector, High is internal rom
  End If
  
  


  
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1.3)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
                Tester.Print LBA
               
                 VidName = "vid_18e3"
                 
                
               
                
                
                If TmpChip = "AU6476DLF20" Then
                
                ' test chip
                      '    CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If LightOff <> 254 And LightOff <> 255 Then
                            UsbSpeedTestResult = GPO_FAIL
                             rv0 = 2
                         End If
          
                End If
                
                
             
'====================================== Assing R/W test switch =====================================
                 
                 
                   TestResult = ""
                   If ChipName = "AU6476DLF20" Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                   End If
              
                    
                   Call MsecDelay(0.1)
                 
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
                 rv0 = CBWTest_New(0, 1, VidName)    ' sd slot
                 
                 If rv0 = 1 Then
                    rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                    
                    If rv0 <> 1 Then
                       rv0 = 2
                       Tester.Print "SD bus width Fail"
                    End If
                 End If
                    
                    
                    
                 
                 
                 
                 If rv0 = 1 And TmpChip = "AU6375HLF22" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                End If
                
                   If Left(TmpChip, 10) = "AU6377ALF2" Then
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                      LBA = TmpLBA
                   End If
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
                
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
              
                
                 rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
               
                ClosePipe
                
                 If rv1 = 1 And rv2 <> 1 Then
                 Tester.Label9.Caption = "SMC FAIL"
                End If
                '============= SMC test begin =======================================
               
              
                 
                          '--- for SMC
                CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
                Call MsecDelay(0.1)
                      
                
                CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
                Call MsecDelay(0.1)
                ClosePipe
                
                
               
                  rv3 = CBWTest_New(2, rv2, VidName)
               
                Call LabelMenu(2, rv3, rv2)
                ClosePipe
                  If rv2 = 1 And rv3 <> 1 Then
                 Tester.Label9.Caption = "XD FAIL"
                End If
                
                
              
                
               '=============== SMC test END ==================================================
               
              
               rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
            
               If rv4 = 1 Then
               rv4 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                If rv4 <> 1 Then
                      rv4 = 2
                      Tester.Print "MS bus width Fail"
                End If
               End If
                   
              
            
               ClosePipe
               Call LabelMenu(3, rv4, rv3)
               
 '======================== light test ======================
              
               
               
               
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
      If rv4 = 1 Then
          CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 high, 7th bit of control 1, and pwr off  for HID mode
          result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
          Call MsecDelay(0.4)
          
          CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
                            
    
          Call MsecDelay(1.2)
          CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
          If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
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
                         GoTo AU6377ALFResult
                     End If
                     
                     '=======================
                     '  key press test, it will return 8 when key up, GPI 6 must do low go hi action
                     '========================
                     
                
                     Do
                          CardResult = DO_WritePort(card, Channel_P1A, &H40) 'GPI6 : bit 6: pull high
                          Sleep (200)
                         CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' GPI6 : bit 6: pull low
                         Sleep (1000)
                       
                         ReturnValue = fnInquiryBtnStatus(DeviceHandle)
                         Tester.Print i; Space(5); "Key press value="; ReturnValue
                         i = i + 1
                     Loop While i < 3 And ReturnValue <> 10
                    ' fnFreeDeviceHandle (DeviceHandle)
                   '  fnFreeDeviceHandle (DeviceHandle)
                     
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
                
AU6377ALFResult:
                        
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
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        
                        End If
                
               
   End Sub

Public Sub AU6476ILF21TestSub()
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
               
   
TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
  
                  
  
         CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  // force into reader mode
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.2)
         CardResult = DO_WritePort(card, Channel_P1A, &H0)
         
         Call MsecDelay(0.3)
                    
        
          
          
   
  
  
 
  
  
  


  
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1.3)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
                Tester.Print LBA
                
                
                VidName = "vid_18e3&pid_9106"
                 '  VidName = "vid_18e3"
                   TestResult = ""
                   
                    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                   
              
                    
                   Call MsecDelay(0.1)
                 
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
                 rv0 = CBWTest_New(0, 1, VidName)     ' sd slot
                 
                 If rv0 = 1 And TmpChip = "AU6375HLF22" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                End If
                
                   If Left(TmpChip, 10) = "AU6377ALF2" Then
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                      LBA = TmpLBA
                   End If
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
                
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
              
                
                 rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
               
                ClosePipe
                
                 If rv1 = 1 And rv2 <> 1 Then
                 Tester.Label9.Caption = "SMC FAIL"
                End If
                '============= SMC test begin =======================================
               
              
                 
                          '--- for SMC
                CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
                Call MsecDelay(0.1)
                      
                
                CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
                Call MsecDelay(0.1)
                ClosePipe
                
                
               
                  rv3 = CBWTest_New(2, rv2, VidName)
               
                Call LabelMenu(2, rv3, rv2)
                ClosePipe
                  If rv2 = 1 And rv3 <> 1 Then
                 Tester.Label9.Caption = "XD FAIL"
                End If
                
                
              
                
               '=============== SMC test END ==================================================
               
              
               rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
            
               ClosePipe
               Call LabelMenu(3, rv4, rv3)
               
 '======================== light test ======================
               
               
               
               
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
      If rv4 = 1 Then
          
          
          'CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
                            
    
         ' Call MsecDelay(1.2)
          CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
          If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
                  UsbSpeedTestResult = GPO_FAIL
                    rv4 = 2
                    Tester.Label9.Caption = "GPO FAIL " & LightOff
           End If
     End If
          
       
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print "LBA="; LBA
                
AU6377ALFResult:
                        
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
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        
                        End If
                
               
                  
                
                
                
                
     '    CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off
     '    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
     '    Call MsecDelay(0.8)
     '    CardResult = DO_WritePort(card, Channel_P1A, &H40)
         
                
                
               
   End Sub
   
Public Sub AU6476ILF24TestSub()
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
               
   
TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
  
                  
  
         CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  // force into reader mode
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.2)
         CardResult = DO_WritePort(card, Channel_P1A, &H0)
         
         Call MsecDelay(0.3)
                    
        
          
          
   
  
  
 
  
  
  


  
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1.3)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
                Tester.Print LBA
                
                
                VidName = "vid_18e3&pid_9106"
                 '  VidName = "vid_18e3"
                   TestResult = ""
                   
                    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                   
              
                    
                   Call MsecDelay(0.1)
                 
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
                 rv0 = CBWTest_New(0, 1, VidName)    ' sd slot
                 
                 If rv0 = 1 Then
                 rv0 = Read_SD_Speed(0, 0, 64, "4Bits")
                 If rv0 <> 1 Then
                     rv0 = 2
                     Tester.Print "SD bus width Fail"
                     End If
                 End If
                 
                 
                 
                 If rv0 = 1 And TmpChip = "AU6375HLF22" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                End If
                
                   If Left(TmpChip, 10) = "AU6377ALF2" Then
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                      LBA = TmpLBA
                   End If
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
                
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
              
                
                 rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
               
                ClosePipe
                
                 If rv1 = 1 And rv2 <> 1 Then
                 Tester.Label9.Caption = "SMC FAIL"
                End If
                '============= SMC test begin =======================================
               
              
                 
                          '--- for SMC
                CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
                Call MsecDelay(0.1)
                      
                
                CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
                Call MsecDelay(0.1)
                ClosePipe
                
                
               
                  rv3 = CBWTest_New(2, rv2, VidName)
               
                Call LabelMenu(2, rv3, rv2)
                ClosePipe
                  If rv2 = 1 And rv3 <> 1 Then
                 Tester.Label9.Caption = "XD FAIL"
                End If
                
                
              
                
               '=============== SMC test END ==================================================
               
              
               rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
               
               If rv4 = 1 Then
               rv4 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                 If rv4 <> 1 Then
                    rv4 = 2
                    Tester.Print "MS bus width fail"
                 End If
               End If
               
            
               ClosePipe
               Call LabelMenu(3, rv4, rv3)
               
 '======================== light test ======================
               
               
               
               
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
      If rv4 = 1 Then
          
          
          'CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
                            
    
         ' Call MsecDelay(1.2)
          CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
          If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
                  UsbSpeedTestResult = GPO_FAIL
                    rv4 = 2
                    Tester.Label9.Caption = "GPO FAIL " & LightOff
           End If
     End If
          
       
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print "LBA="; LBA
                
AU6377ALFResult:
                        
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
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        
                        End If
                
               
                  
                
                
                
                
     '    CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off
     '    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
     '    Call MsecDelay(0.8)
     '    CardResult = DO_WritePort(card, Channel_P1A, &H40)
         
                
                
               
   End Sub
   
Public Sub AU6476VLF26TestSub()
'On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
Dim i As Long
Dim j As Long

   
    TmpChip = ChipName
    '==================================== Switch assign ==========================================
    If PCI7248InitFinish = 0 Then
        PCI7248Exist
    End If
                
    CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  // force into reader mode
    'result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
    Call MsecDelay(0.2)
    CardResult = DO_WritePort(card, Channel_P1A, &H0)
    
    '======================== Begin test ============================================
             
    '//////////////////////////////////////////////////
    '
    '   no card insert
    '
    '/////////////////////////////////////////////////
                
    Dim VidName As String
    LBA = LBA + 1
    
    Tester.Print LBA
    
    VidName = "vid_18e3&pid_9106"
    'Call MsecDelay(0.3)
    WaitDevOn (VidName)
    Call MsecDelay(0.2)
                
    'VidName = "vid_18e3"
    TestResult = ""
                   
    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
    Call MsecDelay(0.1)
                 
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
    rv0 = CBWTest_New(0, 1, VidName)    ' sd slot
                 
    If rv0 = 1 Then
        rv0 = Read_SD_Speed(0, 0, 64, "4Bits")
        If rv0 <> 1 Then
            rv0 = 2
            Tester.Print "SD bus width Fail"
        End If
    End If
                 
    Call LabelMenu(0, rv0, 1)
    ClosePipe
    rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
    
    Call LabelMenu(1, rv1, rv0)
    ClosePipe
                   
    rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
    Call LabelMenu(2, rv2, rv1)
    ClosePipe
                
    If rv1 = 1 And rv2 <> 1 Then
        Tester.Label9.Caption = "SMC FAIL"
    End If
                
    '============= SMC test begin =======================================
    '--- for SMC
    CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
    Call MsecDelay(0.1)
          
    
    CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
    Call MsecDelay(0.1)
    ClosePipe
                
    rv3 = CBWTest_New(2, rv2, VidName)
    
    Call LabelMenu(2, rv3, rv2)
    ClosePipe
    
    If rv2 = 1 And rv3 <> 1 Then
        Tester.Label9.Caption = "XD FAIL"
    End If
             
    '=============== SMC test END ==================================================
              
    rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
    
    If rv4 = 1 Then
        rv4 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
        If rv4 <> 1 Then
            rv4 = 2
            Tester.Print "MS bus width fail"
        End If
    End If
               
    ClosePipe
    Call LabelMenu(3, rv4, rv3)
               
    '======================== light test ======================
               
             
    '=================================================================================
    ' HID mode and reader mode ---> compositive device
    If rv4 = 1 Then
        'CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
        ' Call MsecDelay(1.2)
        CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
        If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
            UsbSpeedTestResult = GPO_FAIL
            rv4 = 2
            Tester.Label9.Caption = "GPO FAIL " & LightOff
        End If
    End If
                
    Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    Tester.Print "LBA="; LBA
                
AU6377ALFResult:
    CardResult = DO_WritePort(card, Channel_P1A, &H80)
    
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
    ElseIf rv4 = WRITE_FAIL Then
        MSWriteFail = MSWriteFail + 1
        TestResult = "MS_WF"
    ElseIf rv4 = READ_FAIL Then
        MSReadFail = MSReadFail + 1
        TestResult = "MS_RF"
    ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
        TestResult = "PASS"
    Else
        TestResult = "Bin2"
    End If
                
End Sub
   
Public Sub AU6476ILF22TestSub()
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
               
   
TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
  
                  
  
         CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  // force into reader mode
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.2)
         CardResult = DO_WritePort(card, Channel_P1A, &H0)
         
         Call MsecDelay(0.3)
                    
        
          
          
   
  
  
 
  
  
  


  
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1.3)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
                Tester.Print LBA
                
                
                VidName = "vid_18e3&pid_9106"
                 '  VidName = "vid_18e3"
                   TestResult = ""
                   
                    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                   
              
                    
                   Call MsecDelay(0.1)
                 
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
                 rv0 = CBWTest_New_SD_Speed(0, 1, VidName, "4Bits")    ' sd slot
                 
                 If rv0 = 1 And TmpChip = "AU6375HLF22" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                End If
                
                   If Left(TmpChip, 10) = "AU6377ALF2" Then
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                      LBA = TmpLBA
                   End If
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
                
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
              
                
                 rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
               
                ClosePipe
                
                 If rv1 = 1 And rv2 <> 1 Then
                 Tester.Label9.Caption = "SMC FAIL"
                End If
                '============= SMC test begin =======================================
               
              
                 
                          '--- for SMC
                CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
                Call MsecDelay(0.1)
                      
                
                CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
                Call MsecDelay(0.1)
                ClosePipe
                
                
               
                  rv3 = CBWTest_New(2, rv2, VidName)
               
                Call LabelMenu(2, rv3, rv2)
                ClosePipe
                  If rv2 = 1 And rv3 <> 1 Then
                 Tester.Label9.Caption = "XD FAIL"
                End If
                
                
              
                
               '=============== SMC test END ==================================================
               
              
               rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
            
               ClosePipe
               Call LabelMenu(3, rv4, rv3)
               
 '======================== light test ======================
               
               
               
               
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
      If rv4 = 1 Then
          
          
          'CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
                            
    
         ' Call MsecDelay(1.2)
          CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
          If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
                  UsbSpeedTestResult = GPO_FAIL
                    rv4 = 2
                    Tester.Label9.Caption = "GPO FAIL " & LightOff
           End If
     End If
          
       
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print "LBA="; LBA
                
AU6377ALFResult:
                        
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
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        
                        End If
                
               
                  
                
                
                
                
     '    CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off
     '    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
     '    Call MsecDelay(0.8)
     '    CardResult = DO_WritePort(card, Channel_P1A, &H40)
         
                
                
               
   End Sub
Public Sub AU6476ELF24TestSub()
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
               
   
TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
  
                  
 If Left(TmpChip, 8) = "AU6378AL" Or Left(TmpChip, 8) = "AU6476EL" Then
 
         CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  // force into reader mode
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.2)
         CardResult = DO_WritePort(card, Channel_P1A, &H0)
         
         Call MsecDelay(0.3)
                    
        
          
          
  End If
  
  
 
  
  
  


  
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1.3)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
                Tester.Print LBA
                
                
                  VidName = "vid_18e3"
          
                   TestResult = ""
                   If ChipName = "AU6476ELF21" Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                   End If
              
                    
                   Call MsecDelay(0.1)
                 
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
                 rv0 = CBWTest_New(0, 1, VidName)    ' sd slot
                 
                 If rv0 = 1 Then
                     rv0 = Read_SD_Speed(0, 0, 64, "4Bits")
                     If rv0 <> 1 Then
                     rv0 = 2
                     Tester.Print "SD bus width Fail"
                     End If
                 End If
                     
                 
                 
                 If rv0 = 1 And TmpChip = "AU6375HLF22" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                End If
                
                   If Left(TmpChip, 10) = "AU6377ALF2" Then
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                      LBA = TmpLBA
                   End If
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
                
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
              
                
                 rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
               
                ClosePipe
                
                 If rv1 = 1 And rv2 <> 1 Then
                 Tester.Label9.Caption = "SMC FAIL"
                End If
                '============= SMC test begin =======================================
               
              
                 
                          '--- for SMC
                CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
                Call MsecDelay(0.1)
                      
                
                CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
                Call MsecDelay(0.1)
                ClosePipe
                
                
               
                  rv3 = CBWTest_New(2, rv2, VidName)
               
                Call LabelMenu(2, rv3, rv2)
                ClosePipe
                  If rv2 = 1 And rv3 <> 1 Then
                 Tester.Label9.Caption = "XD FAIL"
                End If
                
                
              
                
               '=============== SMC test END ==================================================
               
              
               rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
               
               If rv4 = 1 Then
               rv4 = Read_MS_Speed_AU6476(0, 0, 64, "4Bits")
                 If rv4 <> 1 Then
                     rv4 = 2
                     Tester.Print "MS bus width Fail"
                 End If
               End If
               
            
               ClosePipe
               Call LabelMenu(3, rv4, rv3)
               
 '======================== light test ======================
               
               
               
               
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
      If rv4 = 1 Then
          
          
          CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
                            
    
          Call MsecDelay(1.2)
          CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
          If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
                  UsbSpeedTestResult = GPO_FAIL
                    rv4 = 2
                    Tester.Label9.Caption = "GPO FAIL " & LightOff
           End If
     End If
          
       
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print "LBA="; LBA
                
AU6377ALFResult:
                        
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
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        
                        End If
                
               
                  
                
                
                
                
     '    CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off
     '    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
     '    Call MsecDelay(0.8)
     '    CardResult = DO_WritePort(card, Channel_P1A, &H40)
         
                
                
               
   End Sub
  Public Sub AU6476ELF21TestSub()
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
               
   
TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
  
                  
 If Left(TmpChip, 8) = "AU6378AL" Or Left(TmpChip, 8) = "AU6476EL" Then
 
         CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  // force into reader mode
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.2)
         CardResult = DO_WritePort(card, Channel_P1A, &H0)
         
         Call MsecDelay(0.3)
                    
        
          
          
  End If
  
  
 
  
  
  


  
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1.3)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
                Tester.Print LBA
                
                
                  VidName = "vid_18e3"
          
                   TestResult = ""
                   If ChipName = "AU6476ELF21" Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                   End If
              
                    
                   Call MsecDelay(0.1)
                 
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
                 rv0 = CBWTest_New_SD_Speed(0, 1, VidName, "4Bits")    ' sd slot
                 
                 If rv0 = 1 And TmpChip = "AU6375HLF22" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                End If
                
                   If Left(TmpChip, 10) = "AU6377ALF2" Then
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                      LBA = TmpLBA
                   End If
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
                
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
              
                
                 rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
               
                ClosePipe
                
                 If rv1 = 1 And rv2 <> 1 Then
                 Tester.Label9.Caption = "SMC FAIL"
                End If
                '============= SMC test begin =======================================
               
              
                 
                          '--- for SMC
                CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
                Call MsecDelay(0.1)
                      
                
                CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
                Call MsecDelay(0.1)
                ClosePipe
                
                
               
                  rv3 = CBWTest_New(2, rv2, VidName)
               
                Call LabelMenu(2, rv3, rv2)
                ClosePipe
                  If rv2 = 1 And rv3 <> 1 Then
                 Tester.Label9.Caption = "XD FAIL"
                End If
                
                
              
                
               '=============== SMC test END ==================================================
               
              
               rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
            
               ClosePipe
               Call LabelMenu(3, rv4, rv3)
               
 '======================== light test ======================
               
               
               
               
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
      If rv4 = 1 Then
          
          
          CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
                            
    
          Call MsecDelay(1.2)
          CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
          If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
                  UsbSpeedTestResult = GPO_FAIL
                    rv4 = 2
                    Tester.Label9.Caption = "GPO FAIL " & LightOff
           End If
     End If
          
       
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print "LBA="; LBA
                
AU6377ALFResult:
                        
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
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        
                        End If
                
               
                  
                
                
                
                
     '    CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off
     '    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
     '    Call MsecDelay(0.8)
     '    CardResult = DO_WritePort(card, Channel_P1A, &H40)
         
                
                
               
   End Sub
   
Public Sub AU6476ELTestSub()
On Error Resume Next
Dim TmpChip As String
Dim RomSelector As Byte
               
   
TmpChip = ChipName
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
  
                  
 If Left(TmpChip, 8) = "AU6378AL" Or Left(TmpChip, 8) = "AU6476EL" Then
 
         CardResult = DO_WritePort(card, Channel_P1A, &H80) ' pull gpi6 low, and pwr off  // force into reader mode
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.2)
         CardResult = DO_WritePort(card, Channel_P1A, &H0)
         
         Call MsecDelay(0.3)
                    
        
          
          
  End If
  
  
 
  
  
  


  
  
'======================== Begin test ============================================
                  
                Call MsecDelay(1.3)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
                Tester.Print LBA
                
                
                  VidName = "vid_18e3"
          
                   TestResult = ""
                   If ChipName = "AU6476ELF20" Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                   End If
              
                    
                   Call MsecDelay(0.1)
                 
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
                 rv0 = CBWTest_New(0, 1, VidName)     ' sd slot
                 
                 If rv0 = 1 And TmpChip = "AU6375HLF22" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                End If
                
                   If Left(TmpChip, 10) = "AU6377ALF2" Then
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6377(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                      LBA = TmpLBA
                   End If
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 rv1 = CBWTest_New(1, rv0, VidName)    ' cf slot
                
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
              
                
                 rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
               
                ClosePipe
                
                 If rv1 = 1 And rv2 <> 1 Then
                 Tester.Label9.Caption = "SMC FAIL"
                End If
                '============= SMC test begin =======================================
               
              
                 
                          '--- for SMC
                CardResult = DO_WritePort(card, Channel_P1A, &HC)   ' 0110 0100
                Call MsecDelay(0.1)
                      
                
                CardResult = DO_WritePort(card, Channel_P1A, &H4)  ' 0110 0100 SMC high
                Call MsecDelay(0.1)
                ClosePipe
                
                
               
                  rv3 = CBWTest_New(2, rv2, VidName)
               
                Call LabelMenu(2, rv3, rv2)
                ClosePipe
                  If rv2 = 1 And rv3 <> 1 Then
                 Tester.Label9.Caption = "XD FAIL"
                End If
                
                
              
                
               '=============== SMC test END ==================================================
               
              
               rv4 = CBWTest_New(3, rv3, VidName)   'MS card test
            
               ClosePipe
               Call LabelMenu(3, rv4, rv3)
               
 '======================== light test ======================
               
               
               
               
 '=================================================================================
 ' HID mode and reader mode ---> compositive device
      If rv4 = 1 Then
          
          
          CardResult = DO_WritePort(card, Channel_P1A, &H7D) ' HID mode
                            
    
          Call MsecDelay(1.2)
          CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
          If rv4 = 1 And (LightOff <> 252 And LightOff <> 254) Then
                  UsbSpeedTestResult = GPO_FAIL
                    rv4 = 2
                    Tester.Label9.Caption = "GPO FAIL " & LightOff
           End If
     End If
          
       
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print "LBA="; LBA
                
AU6377ALFResult:
                        
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
                         ElseIf rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        
                        End If
                
               
                  
                
                
                
                
     '    CardResult = DO_WritePort(card, Channel_P1A, &HC0) ' pull gpi6 low, and pwr off
     '    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
     '    Call MsecDelay(0.8)
     '    CardResult = DO_WritePort(card, Channel_P1A, &H40)
         
                
                
               
   End Sub
