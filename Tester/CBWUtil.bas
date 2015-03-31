Attribute VB_Name = "CBWUtil"
Option Explicit
Dim TmpLBA As Long
Dim i As Integer

  Public Sub MultiSlotTestAU6376ALO12()
' add XD MS data pin bonding error sorting
Dim TmpChip As String
Dim RomSelector As Byte
               
  Call PowerSet2(0, "3.25", "0.5", 1, "3.25", "0.5", 1)
  If ChipName = "AU6370GLF20" Then
      ChipName = "AU6370DLF20"
  End If
                
                ' open power
 If ChipName = "AU6377ALF24" Or ChipName = "AU6377ALF25" Then
     TmpChip = ChipName
     ChipName = "AU6376"
 End If
                
                
            '    PowerSet (1) ' for 3.3V , 2.5 V
 If ChipName = "AU6370DLF20" Or ChipName = "AU6378ALF20" Then
     TmpChip = ChipName
     ChipName = "AU6376"
 End If
            
                'GPIO control setting
If ChipName = "AU6370BL" Or InStr(ChipName, "AU6375HL") <> 0 Or ChipName = "AU6375CL" Or ChipName = "AU6377ALF21" Or ChipName = "AU6377ALS10" Then
     TmpChip = ChipName
     ChipName = "AU6376"
End If
                
If ChipName = "AU6376ELF22" Or ChipName = "AU6376ILF20" Then
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
                  
 
                  
 If TmpChip = "AU6378ALF20" Then
                    
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
                If TmpChip = "AU6377ALF25" Then
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
                
                
                If TmpChip = "AU6378ALF20" Then
                
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
                   '
                If TestResult = "PASS" Then
                  TestResult = ""
                  
                   
                  
                    CardResult = DO_WritePort(card, Channel_P1A, &H1A)  ' 0110 0100  only CF open
                   
                   
                
                   
                    
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
                 
                
                 rv0 = CBWTest_New(1, 1, VidName)     ' Test CF at 1st slot
                 
                 If rv0 = 1 And Left(TmpChip, 10) = "AU6375HLF2" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                    
                    ' AU6375 ram unstable
                    
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 5
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6375(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                    
                    
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
                Call LabelMenu(1, rv0, 1)
                
                
                
                
                
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H10)  ' 0110 0100  SD open,XD open
                   
                ClosePipe
                 rv1 = CBWTest_New(0, rv0, VidName)    ' SD slot
            
                Call LabelMenu(0, rv1, rv0)
                ClosePipe
              
                rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
                
                '============= SMC test begin =======================================
               
                If rv2 = 1 And TmpChip = "AU6378ALF20" Then         '--- for SMC
                
                CardResult = DO_WritePort(card, Channel_P1A, &H18)  ' 0110 0100
                Call MsecDelay(0.5)
                ClosePipe
                rv2 = CBWTest_New(2, rv2, VidName)
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
                 End If
                 
              If rv2 = 1 And Left(TmpChip, 10) = "AU6377ALF2" And TmpChip <> "AU6377ALF21" Then           '--- for SMC
                
                CardResult = DO_WritePort(card, Channel_P1A, &H8 + RomSelector) ' 0110 0100
                Call MsecDelay(0.5)
                ClosePipe
                rv2 = CBWTest_New(2, rv2, VidName)
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
              End If
                
                
               If rv2 = 1 And (TmpChip = "AU6376JLF20") Then      '--- for SMC
                
                  CardResult = DO_WritePort(card, Channel_P1A, &H18)   ' 0110 0100
                  Call MsecDelay(0.5)
                  CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                  Call MsecDelay(0.5)
                 ClosePipe
                 rv2 = CBWTest_New(2, rv2, VidName)
                 Call LabelMenu(2, rv2, rv1)
                 ClosePipe
               End If
               
               
                CardResult = DO_WritePort(card, Channel_P1A, &H18)   ' 0110 0100
                  Call MsecDelay(0.5)
                  CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                  Call MsecDelay(0.5)
                 ClosePipe
                 rv2 = CBWTest_New(2, rv2, VidName)
                 Call LabelMenu(2, rv2, rv1)
                 ClosePipe
               
              
          
               '=============== SMC test END ==================================================
               
               rv3 = CBWTest_New(3, rv2, VidName)  ' MS test
               ClosePipe
               Call LabelMenu(3, rv3, rv2)
             '========================================================
             
                  CardResult = DO_WritePort(card, Channel_P1A, &H18)   ' 0110 0100
                  Call MsecDelay(0.5)
                  CardResult = DO_WritePort(card, Channel_P1A, &H10)   ' 0110 0100
                  Call MsecDelay(0.5)
           
                  rv2 = CBWTest_New(2, rv2, VidName)
                
                  Call LabelMenu(2, rv2, rv1)
                  ClosePipe
           
             
               If TmpChip = "AU6375HLF21" Then
               
                 If rv0 = 1 Then
                   
                    ClosePipe
                     rv0 = CBWTest_New_AU6375IncPattern(0, 1, VidName)
                     Call LabelMenu(0, rv0, 1)
                     ClosePipe
                 End If
                
                End If
                
                 
                 
                If Left(TmpChip, 10) = "AU6377ALF2" Then
                
                ' test chip
                         ClosePipe
                         rv4 = CBWTest_New(4, rv3, VidName)   'MMC test
                          Call LabelMenu(10, rv4, rv3)
                          ClosePipe
          
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If LightOff <> 127 Then
                            UsbSpeedTestResult = GPO_FAIL
                            rv4 = 2
                         End If
          
                End If
                 
                If TmpChip = "AU6378ALF20" Then
                
             
                         ClosePipe
                         rv4 = CBWTest_New(4, rv3, VidName)
                          Call LabelMenu(10, rv4, rv3)
                          ClosePipe
          
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If rv4 = 1 And LightOff <> 252 Then
                            UsbSpeedTestResult = GPO_FAIL
                            rv4 = 2
                         End If
          
                End If
                 
                 
                    
                  If ChipName = "AU6376" And TmpChip = "AU6370DLF20" Then
                  Call MsecDelay(0.1)
                  rv4 = 1
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 If LightOff <> 254 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                 End If
                     Call LabelMenu(3, rv4, rv3)
                     
                        
                 End If
                 
                 
                 If ChipName = "AU6368A1" Then
                 Call MsecDelay(0.1)
                  rv4 = 1
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 If LightOff <> 192 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                 End If
                     Call LabelMenu(3, rv4, rv3)
                     
                       CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 
                 End If
                 
                   If ChipName = "AU6376" And (Left(TmpChip, 10) <> "AU6377ALF2" And TmpChip <> "AU6378ALF20") Then
                   Call MsecDelay(0.1)
                   rv4 = 1
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                If TmpChip = "AU6370DLF20" Or TmpChip = "AU6376JLF20" Then
                               
                                   If LightOff <> 254 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                     
                                  rv4 = 2
                                 End If
                                Else
                                  If LightOff <> 252 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                
                                 rv4 = 2
                                 End If
                             End If
                     Call LabelMenu(3, rv4, rv3)
                     
                        
                 End If
                 
                If ChipName = "AU6368A" Then
                    If rv3 = 1 Then
                           CardResult = DO_WritePort(card, Channel_P1A, &H74)  ' 0111 0100
                           Call MsecDelay(0.1)
                           CardResult = DO_WritePort(card, Channel_P1A, &H54)  ' 0101 0100
                           Call MsecDelay(0.1)
                           rv4 = CBWTest_New(3, rv3, VidName)
                             ClosePipe
                           If rv4 = 1 Then
                                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 If LightOff <> 132 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                    rv4 = 2
                                 End If
                             End If
                         Else
                         rv4 = 4
                         End If
                         Call LabelMenu(3, rv4, rv3)
                 End If
                
                  OpenPipe
                  If rv4 = 1 Then
                   
                    If rv4 = 1 Then
                     rv5 = SetOverCurrent(rv4)
                        If rv5 <> 1 Then
                          rv5 = 2
                        End If
                     End If
                     
                   If rv5 = 1 Then
                        rv5 = Read_OverCurrent(0, 0, 64)
                        If rv5 <> 1 Then
                        rv5 = 2
                        End If
                   End If
                   
                 End If
                    ClosePipe
                    Call LabelMenu(51, rv5, rv4)
                    If rv5 <> 1 Then
                    Tester.Label9.Caption = "Over Current fail"
                    Tester.Label2.Caption = "Over Current fail---"
                    End If
                   
                Tester.Print rv0, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv5, " \\OverCurret :0 Fail, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
               
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
                        ElseIf rv2 = WRITE_FAIL Or rv3 = WRITE_FAIL Or rv4 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv2 = READ_FAIL Or rv3 = READ_FAIL Or rv4 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                         ElseIf rv5 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv5 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                            
                            
                        ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
                
               
                  
                End If
                  
                CardResult = DO_WritePort(card, Channel_P1A, &H1)
                  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    
   End Sub
   
  Public Sub MultiSlotTestAU6376ALO13()
' add XD MS data pin bonding error sorting
' 9870827 : set true CF overcurretn fai at Bin4
'

Dim TmpChip As String
Dim RomSelector As Byte
               
  Call PowerSet2(0, "3.25", "0.5", 1, "3.25", "0.5", 1)
  If ChipName = "AU6370GLF20" Then
      ChipName = "AU6370DLF20"
  End If
                
                ' open power
 If ChipName = "AU6377ALF24" Or ChipName = "AU6377ALF25" Then
     TmpChip = ChipName
     ChipName = "AU6376"
 End If
                
                
            '    PowerSet (1) ' for 3.3V , 2.5 V
 If ChipName = "AU6370DLF20" Or ChipName = "AU6378ALF20" Then
     TmpChip = ChipName
     ChipName = "AU6376"
 End If
            
                'GPIO control setting
If ChipName = "AU6370BL" Or InStr(ChipName, "AU6375HL") <> 0 Or ChipName = "AU6375CL" Or ChipName = "AU6377ALF21" Or ChipName = "AU6377ALS10" Then
     TmpChip = ChipName
     ChipName = "AU6376"
End If
                
If ChipName = "AU6376ELF22" Or ChipName = "AU6376ILF20" Then
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
                  
 
                  
 If TmpChip = "AU6378ALF20" Then
                    
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
                If TmpChip = "AU6377ALF25" Then
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
                
                
                If TmpChip = "AU6378ALF20" Then
                
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
                   '
                If TestResult = "PASS" Then
                  TestResult = ""
                  
                   
                  
                    CardResult = DO_WritePort(card, Channel_P1A, &H1A)  ' 0110 0100  only CF open
                   
                   
                
                   
                    
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
                 
                
                 rv0 = CBWTest_New(1, 1, VidName)     ' Test CF at 1st slot
                 
                 If rv0 = 1 And Left(TmpChip, 10) = "AU6375HLF2" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                    
                    ' AU6375 ram unstable
                    
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 5
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6375(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                    
                    
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
                Call LabelMenu(1, rv0, 1)
                
                
                
                
                
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H10)  ' 0110 0100  SD open,XD open
                   
                ClosePipe
                 rv1 = CBWTest_New(0, rv0, VidName)    ' SD slot
            
                Call LabelMenu(0, rv1, rv0)
                ClosePipe
              
                rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
                
                '============= SMC test begin =======================================
               
                If rv2 = 1 And TmpChip = "AU6378ALF20" Then         '--- for SMC
                
                CardResult = DO_WritePort(card, Channel_P1A, &H18)  ' 0110 0100
                Call MsecDelay(0.5)
                ClosePipe
                rv2 = CBWTest_New(2, rv2, VidName)
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
                 End If
                 
              If rv2 = 1 And Left(TmpChip, 10) = "AU6377ALF2" And TmpChip <> "AU6377ALF21" Then           '--- for SMC
                
                CardResult = DO_WritePort(card, Channel_P1A, &H8 + RomSelector) ' 0110 0100
                Call MsecDelay(0.5)
                ClosePipe
                rv2 = CBWTest_New(2, rv2, VidName)
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
              End If
                
                
               If rv2 = 1 And (TmpChip = "AU6376JLF20") Then      '--- for SMC
                
                  CardResult = DO_WritePort(card, Channel_P1A, &H18)   ' 0110 0100
                  Call MsecDelay(0.5)
                  CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                  Call MsecDelay(0.5)
                 ClosePipe
                 rv2 = CBWTest_New(2, rv2, VidName)
                 Call LabelMenu(2, rv2, rv1)
                 ClosePipe
               End If
               
               
                CardResult = DO_WritePort(card, Channel_P1A, &H18)   ' 0110 0100
                  Call MsecDelay(0.5)
                  CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                  Call MsecDelay(0.5)
                 ClosePipe
                 rv2 = CBWTest_New(2, rv2, VidName)
                 Call LabelMenu(2, rv2, rv1)
                 ClosePipe
               
              
          
               '=============== SMC test END ==================================================
               
               rv3 = CBWTest_New(3, rv2, VidName)  ' MS test
               ClosePipe
               Call LabelMenu(3, rv3, rv2)
             '========================================================
             
                  CardResult = DO_WritePort(card, Channel_P1A, &H18)   ' 0110 0100
                  Call MsecDelay(0.5)
                  CardResult = DO_WritePort(card, Channel_P1A, &H10)   ' 0110 0100
                  Call MsecDelay(0.5)
           
                  rv2 = CBWTest_New(2, rv2, VidName)
                
                  Call LabelMenu(2, rv2, rv1)
                  ClosePipe
           
             
               If TmpChip = "AU6375HLF21" Then
               
                 If rv0 = 1 Then
                   
                    ClosePipe
                     rv0 = CBWTest_New_AU6375IncPattern(0, 1, VidName)
                     Call LabelMenu(0, rv0, 1)
                     ClosePipe
                 End If
                
                End If
                
                 
                 
                If Left(TmpChip, 10) = "AU6377ALF2" Then
                
                ' test chip
                         ClosePipe
                         rv4 = CBWTest_New(4, rv3, VidName)   'MMC test
                          Call LabelMenu(10, rv4, rv3)
                          ClosePipe
          
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If LightOff <> 127 Then
                            UsbSpeedTestResult = GPO_FAIL
                            rv4 = 2
                         End If
          
                End If
                 
                If TmpChip = "AU6378ALF20" Then
                
             
                         ClosePipe
                         rv4 = CBWTest_New(4, rv3, VidName)
                          Call LabelMenu(10, rv4, rv3)
                          ClosePipe
          
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If rv4 = 1 And LightOff <> 252 Then
                            UsbSpeedTestResult = GPO_FAIL
                            rv4 = 2
                         End If
          
                End If
                 
                 
                    
                  If ChipName = "AU6376" And TmpChip = "AU6370DLF20" Then
                  Call MsecDelay(0.1)
                  rv4 = 1
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 If LightOff <> 254 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                 End If
                     Call LabelMenu(3, rv4, rv3)
                     
                        
                 End If
                 
                 
                 If ChipName = "AU6368A1" Then
                 Call MsecDelay(0.1)
                  rv4 = 1
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 If LightOff <> 192 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                 End If
                     Call LabelMenu(3, rv4, rv3)
                     
                       CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 
                 End If
                 
                   If ChipName = "AU6376" And (Left(TmpChip, 10) <> "AU6377ALF2" And TmpChip <> "AU6378ALF20") Then
                   Call MsecDelay(0.1)
                   rv4 = 1
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                If TmpChip = "AU6370DLF20" Or TmpChip = "AU6376JLF20" Then
                               
                                   If LightOff <> 254 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                     
                                  rv4 = 2
                                 End If
                                Else
                                  If LightOff <> 252 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                
                                 rv4 = 2
                                 End If
                             End If
                     Call LabelMenu(3, rv4, rv3)
                     
                        
                 End If
                 
                If ChipName = "AU6368A" Then
                    If rv3 = 1 Then
                           CardResult = DO_WritePort(card, Channel_P1A, &H74)  ' 0111 0100
                           Call MsecDelay(0.1)
                           CardResult = DO_WritePort(card, Channel_P1A, &H54)  ' 0101 0100
                           Call MsecDelay(0.1)
                           rv4 = CBWTest_New(3, rv3, VidName)
                             ClosePipe
                           If rv4 = 1 Then
                                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 If LightOff <> 132 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                    rv4 = 2
                                 End If
                             End If
                         Else
                         rv4 = 4
                         End If
                         Call LabelMenu(3, rv4, rv3)
                 End If
                
                  OpenPipe
                  If rv4 = 1 Then
                   
                    If rv4 = 1 Then
                     rv4 = SetOverCurrent(rv4)
                        If rv4 <> 1 Then
                          rv4 = 2
                        End If
                     End If
                     
                   If rv4 = 1 Then
                        rv5 = Read_OverCurrent(0, 0, 64)
                      
                   End If
                   
                 End If
                    ClosePipe
                    Call LabelMenu(51, rv5, rv4)
                     
                    If rv5 = 3 Then
                    Tester.Label9.Caption = "Over Current fail"
                    Tester.Label2.Caption = "Over Current fail---"
                    End If
                   
                Tester.Print rv0, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv5, " \\OverCurret :0 Fail, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
               
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
                        ElseIf rv2 = WRITE_FAIL Or rv3 = WRITE_FAIL Or rv4 = WRITE_FAIL Then
                           
                             MSWriteFail = MSWriteFail + 1
                             TestResult = "MS_WF"
                            
                        ElseIf rv5 = 0 Or rv2 = READ_FAIL Or rv3 = READ_FAIL Or rv4 = READ_FAIL Then
                              MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                         
                           
                        ElseIf rv5 = 3 Then
                             XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                            
                            
                        ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
                
               
                  
                End If
                  
                CardResult = DO_WritePort(card, Channel_P1A, &H1)
                  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    
   End Sub
   
     Public Sub MultiSlotTestAU6376ALO14()
' add XD MS data pin bonding error sorting
' 9870827 : set true CF overcurretn fai at Bin4
'

Dim TmpChip As String
Dim RomSelector As Byte
               
  Call PowerSet2(0, "3.25", "0.5", 1, "3.25", "0.5", 1)
  If ChipName = "AU6370GLF20" Then
      ChipName = "AU6370DLF20"
  End If
                
                ' open power
 If ChipName = "AU6377ALF24" Or ChipName = "AU6377ALF25" Then
     TmpChip = ChipName
     ChipName = "AU6376"
 End If
                
                
            '    PowerSet (1) ' for 3.3V , 2.5 V
 If ChipName = "AU6370DLF20" Or ChipName = "AU6378ALF20" Then
     TmpChip = ChipName
     ChipName = "AU6376"
 End If
            
                'GPIO control setting
If ChipName = "AU6370BL" Or InStr(ChipName, "AU6375HL") <> 0 Or ChipName = "AU6375CL" Or ChipName = "AU6377ALF21" Or ChipName = "AU6377ALS10" Then
     TmpChip = ChipName
     ChipName = "AU6376"
End If
                
If ChipName = "AU6376ELF22" Or ChipName = "AU6376ILF20" Then
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
                  
 
                  
 If TmpChip = "AU6378ALF20" Then
                    
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
                If TmpChip = "AU6377ALF25" Then
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
                
                
                If TmpChip = "AU6378ALF20" Then
                
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
                   '
                If TestResult = "PASS" Then
                  TestResult = ""
                  
                   
                  
                    CardResult = DO_WritePort(card, Channel_P1A, &H1A)  ' 0110 0100  only CF open
                   
                   
                
                   
                    
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
                 
                
                 rv0 = CBWTest_NewCFOverCurrentRW(1, 1, VidName)     ' Test CF at 1st slot
                 
                 If rv0 = 5 Then
                    rv5 = 5
                    GoTo OverCurrentLabel
                 End If
                 
                 If rv0 = 1 And Left(TmpChip, 10) = "AU6375HLF2" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                    
                    ' AU6375 ram unstable
                    
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 5
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6375(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                    
                    
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
                Call LabelMenu(1, rv0, 1)
                
                
                
                
                
                
                
                 CardResult = DO_WritePort(card, Channel_P1A, &H10)  ' 0110 0100  SD open,XD open
                   
                ClosePipe
                 rv1 = CBWTest_New(0, rv0, VidName)    ' SD slot
            
                Call LabelMenu(0, rv1, rv0)
                ClosePipe
              
                rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
                
                '============= SMC test begin =======================================
               
                If rv2 = 1 And TmpChip = "AU6378ALF20" Then         '--- for SMC
                
                CardResult = DO_WritePort(card, Channel_P1A, &H18)  ' 0110 0100
                Call MsecDelay(0.5)
                ClosePipe
                rv2 = CBWTest_New(2, rv2, VidName)
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
                 End If
                 
              If rv2 = 1 And Left(TmpChip, 10) = "AU6377ALF2" And TmpChip <> "AU6377ALF21" Then           '--- for SMC
                
                CardResult = DO_WritePort(card, Channel_P1A, &H8 + RomSelector) ' 0110 0100
                Call MsecDelay(0.5)
                ClosePipe
                rv2 = CBWTest_New(2, rv2, VidName)
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
              End If
                
                
               If rv2 = 1 And (TmpChip = "AU6376JLF20") Then      '--- for SMC
                
                  CardResult = DO_WritePort(card, Channel_P1A, &H18)   ' 0110 0100
                  Call MsecDelay(0.5)
                  CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                  Call MsecDelay(0.5)
                 ClosePipe
                 rv2 = CBWTest_New(2, rv2, VidName)
                 Call LabelMenu(2, rv2, rv1)
                 ClosePipe
               End If
               
               
                CardResult = DO_WritePort(card, Channel_P1A, &H18)   ' 0110 0100
                  Call MsecDelay(0.5)
                  CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                  Call MsecDelay(0.5)
                 ClosePipe
                 rv2 = CBWTest_New(2, rv2, VidName)
                 Call LabelMenu(2, rv2, rv1)
                 ClosePipe
               
              
          
               '=============== SMC test END ==================================================
               
               rv3 = CBWTest_New(3, rv2, VidName)  ' MS test
               ClosePipe
               Call LabelMenu(3, rv3, rv2)
             '========================================================
             
                  CardResult = DO_WritePort(card, Channel_P1A, &H18)   ' 0110 0100
                  Call MsecDelay(0.5)
                  CardResult = DO_WritePort(card, Channel_P1A, &H10)   ' 0110 0100
                  Call MsecDelay(0.5)
           
                  rv2 = CBWTest_New(2, rv2, VidName)
                
                  Call LabelMenu(2, rv2, rv1)
                  ClosePipe
           
             
               If TmpChip = "AU6375HLF21" Then
               
                 If rv0 = 1 Then
                   
                    ClosePipe
                     rv0 = CBWTest_New_AU6375IncPattern(0, 1, VidName)
                     Call LabelMenu(0, rv0, 1)
                     ClosePipe
                 End If
                
                End If
                
                 
                 
                If Left(TmpChip, 10) = "AU6377ALF2" Then
                
                ' test chip
                         ClosePipe
                         rv4 = CBWTest_New(4, rv3, VidName)   'MMC test
                          Call LabelMenu(10, rv4, rv3)
                          ClosePipe
          
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If LightOff <> 127 Then
                            UsbSpeedTestResult = GPO_FAIL
                            rv4 = 2
                         End If
          
                End If
                 
                If TmpChip = "AU6378ALF20" Then
                
             
                         ClosePipe
                         rv4 = CBWTest_New(4, rv3, VidName)
                          Call LabelMenu(10, rv4, rv3)
                          ClosePipe
          
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If rv4 = 1 And LightOff <> 252 Then
                            UsbSpeedTestResult = GPO_FAIL
                            rv4 = 2
                         End If
          
                End If
                 
                 
                    
                  If ChipName = "AU6376" And TmpChip = "AU6370DLF20" Then
                  Call MsecDelay(0.1)
                  rv4 = 1
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 If LightOff <> 254 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                 End If
                     Call LabelMenu(3, rv4, rv3)
                     
                        
                 End If
                 
                 
                 If ChipName = "AU6368A1" Then
                 Call MsecDelay(0.1)
                  rv4 = 1
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 If LightOff <> 192 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                 End If
                     Call LabelMenu(3, rv4, rv3)
                     
                       CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 
                 End If
                 
                   If ChipName = "AU6376" And (Left(TmpChip, 10) <> "AU6377ALF2" And TmpChip <> "AU6378ALF20") Then
                   Call MsecDelay(0.1)
                   rv4 = 1
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                If TmpChip = "AU6370DLF20" Or TmpChip = "AU6376JLF20" Then
                               
                                   If LightOff <> 254 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                     
                                  rv4 = 2
                                 End If
                                Else
                                  If LightOff <> 252 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                
                                 rv4 = 2
                                 End If
                             End If
                     Call LabelMenu(3, rv4, rv3)
                     
                        
                 End If
                 
                If ChipName = "AU6368A" Then
                    If rv3 = 1 Then
                           CardResult = DO_WritePort(card, Channel_P1A, &H74)  ' 0111 0100
                           Call MsecDelay(0.1)
                           CardResult = DO_WritePort(card, Channel_P1A, &H54)  ' 0101 0100
                           Call MsecDelay(0.1)
                           rv4 = CBWTest_New(3, rv3, VidName)
                             ClosePipe
                           If rv4 = 1 Then
                                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 If LightOff <> 132 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                    rv4 = 2
                                 End If
                             End If
                         Else
                         rv4 = 4
                         End If
                         Call LabelMenu(3, rv4, rv3)
                 End If
                
                  
                  
OverCurrentLabel:
                   
                     
                    If rv5 = 5 Then
                     Call LabelMenu(51, rv5, rv4)
                    Tester.Label9.Caption = "Over Current fail"
                    Tester.Label2.Caption = "Over Current fail---"
                    End If
                   
                Tester.Print rv0, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv5, " \\OverCurret :5 Fail, 0 pass ,2 write fail, 3 read fail, 4 preious slot fail"
               
                Tester.Print "LBA="; LBA
    
                 
                
AU6377ALFResult:

                        If rv5 = 5 Then
                             XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                            
                            
                        
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
 
 Public Sub MultiSlotTestAU6376ALOT1()
' add XD MS data pin bonding error sorting
 
' 980827  : set overcurrent flow , is test unit ready only
Tester.Print "3.25 V verify"
Dim TmpChip As String
Dim RomSelector As Byte
               
  Call PowerSet2(0, "3.25", "0.5", 1, "3.25", "0.5", 1)
  If ChipName = "AU6370GLF20" Then
      ChipName = "AU6370DLF20"
  End If
                
                ' open power
 If ChipName = "AU6377ALF24" Or ChipName = "AU6377ALF25" Then
     TmpChip = ChipName
     ChipName = "AU6376"
 End If
                
                
            '    PowerSet (1) ' for 3.3V , 2.5 V
 If ChipName = "AU6370DLF20" Or ChipName = "AU6378ALF20" Then
     TmpChip = ChipName
     ChipName = "AU6376"
 End If
            
                'GPIO control setting
If ChipName = "AU6370BL" Or InStr(ChipName, "AU6375HL") <> 0 Or ChipName = "AU6375CL" Or ChipName = "AU6377ALF21" Or ChipName = "AU6377ALS10" Then
     TmpChip = ChipName
     ChipName = "AU6376"
End If
                
If ChipName = "AU6376ELF22" Or ChipName = "AU6376ILF20" Then
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
                  
 
                  
 If TmpChip = "AU6378ALF20" Then
                    
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
                If TmpChip = "AU6377ALF25" Then
                  VidName = "vid_1984"
                Else
                 VidName = "vid_058f"
                End If
                
              
               
 '================================= Test light off =============================
                
            
             
                  TestResult = ""
                  
                   
                  
                    CardResult = DO_WritePort(card, Channel_P1A, &H3A)  ' 0110 0100  only CF open
                   
                   
                
                   
                    
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
                
   
                
                 rv0 = CBWTest_NewOverCurrent(1, 1, VidName)     ' Test CF at 1st slot
                 
       
                 Call LabelMenu(0, rv0, 1)
                     
                    If rv0 = 3 Then
                    Tester.Label9.Caption = "Over Current fail"
                    Tester.Label2.Caption = "Over Current fail---"
                    End If
               
                   
                Tester.Print rv0, " \\CF :0 Unknow device, 1 pass ,2   fail, 3 over currnt, 4 preious slot fail"
                
               
                Tester.Print "LBA="; LBA
    
                 
                
AU6377ALFResult:
                        
                        If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                        ElseIf rv0 = 2 Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
                        ElseIf rv0 = 3 Then
                        
                             XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                            
                            
                        ElseIf rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
                
               
                  
               
                  
                CardResult = DO_WritePort(card, Channel_P1A, &H1)
                  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    
   End Sub
   
 Public Sub MultiSlotTestAU6376ALOT2()
' add XD MS data pin bonding error sorting
 
' 980827  : set overcurrent flow , is test unit ready only
Tester.Print "3.20 V verify"
Dim TmpChip As String
Dim RomSelector As Byte
               
  Call PowerSet2(0, "3.20", "0.5", 1, "3.20", "0.5", 1)
  If ChipName = "AU6370GLF20" Then
      ChipName = "AU6370DLF20"
  End If
                
                ' open power
 If ChipName = "AU6377ALF24" Or ChipName = "AU6377ALF25" Then
     TmpChip = ChipName
     ChipName = "AU6376"
 End If
                
                
            '    PowerSet (1) ' for 3.3V , 2.5 V
 If ChipName = "AU6370DLF20" Or ChipName = "AU6378ALF20" Then
     TmpChip = ChipName
     ChipName = "AU6376"
 End If
            
                'GPIO control setting
If ChipName = "AU6370BL" Or InStr(ChipName, "AU6375HL") <> 0 Or ChipName = "AU6375CL" Or ChipName = "AU6377ALF21" Or ChipName = "AU6377ALS10" Then
     TmpChip = ChipName
     ChipName = "AU6376"
End If
                
If ChipName = "AU6376ELF22" Or ChipName = "AU6376ILF20" Then
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
                  
 
                  
 If TmpChip = "AU6378ALF20" Then
                    
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
                If TmpChip = "AU6377ALF25" Then
                  VidName = "vid_1984"
                Else
                 VidName = "vid_058f"
                End If
                
              
               
 '================================= Test light off =============================
                
            
             
                  TestResult = ""
                  
                   
                  
                    CardResult = DO_WritePort(card, Channel_P1A, &H1A)  ' 0110 0100  only CF open
                   
                   
                
                   
                    
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
                
   
                
                 rv0 = CBWTest_NewOverCurrent(1, 1, VidName)     ' Test CF at 1st slot
                 
       
                 Call LabelMenu(0, rv0, 1)
                     
                    If rv0 = 3 Then
                    Tester.Label9.Caption = "Over Current fail"
                    Tester.Label2.Caption = "Over Current fail---"
                    End If
               
                   
                Tester.Print rv0, " \\CF :0 Unknow device, 1 pass ,2   fail, 3 over currnt, 4 preious slot fail"
                
               
                Tester.Print "LBA="; LBA
    
                 
                
AU6377ALFResult:
                        
                        If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                        ElseIf rv0 = 2 Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
                        ElseIf rv0 = 3 Then
                        
                             XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                            
                            
                        ElseIf rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
                
               
                  
               
                  
                CardResult = DO_WritePort(card, Channel_P1A, &H1)
                  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    
   End Sub
   
    Public Sub MultiSlotTestAU6376ALOT3()
' add XD MS data pin bonding error sorting
 
' 980827  : set overcurrent flow , is test unit ready only
Tester.Print "3.15 V verify"
Dim TmpChip As String
Dim RomSelector As Byte
               
  Call PowerSet2(0, "3.15", "0.5", 1, "3.15", "0.5", 1)
  If ChipName = "AU6370GLF20" Then
      ChipName = "AU6370DLF20"
  End If
                
                ' open power
 If ChipName = "AU6377ALF24" Or ChipName = "AU6377ALF25" Then
     TmpChip = ChipName
     ChipName = "AU6376"
 End If
                
                
            '    PowerSet (1) ' for 3.3V , 2.5 V
 If ChipName = "AU6370DLF20" Or ChipName = "AU6378ALF20" Then
     TmpChip = ChipName
     ChipName = "AU6376"
 End If
            
                'GPIO control setting
If ChipName = "AU6370BL" Or InStr(ChipName, "AU6375HL") <> 0 Or ChipName = "AU6375CL" Or ChipName = "AU6377ALF21" Or ChipName = "AU6377ALS10" Then
     TmpChip = ChipName
     ChipName = "AU6376"
End If
                
If ChipName = "AU6376ELF22" Or ChipName = "AU6376ILF20" Then
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
                  
 
                  
 If TmpChip = "AU6378ALF20" Then
                    
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
                If TmpChip = "AU6377ALF25" Then
                  VidName = "vid_1984"
                Else
                 VidName = "vid_058f"
                End If
                
              
               
 '================================= Test light off =============================
                
            
             
                  TestResult = ""
                  
                   
                  
                    CardResult = DO_WritePort(card, Channel_P1A, &H1A)  ' 0110 0100  only CF open
                   
                   
                
                   
                    
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
                
   
                
                 rv0 = CBWTest_NewOverCurrent(1, 1, VidName)     ' Test CF at 1st slot
                 
       
                 Call LabelMenu(0, rv0, 1)
                     
                    If rv0 = 3 Then
                    Tester.Label9.Caption = "Over Current fail"
                    Tester.Label2.Caption = "Over Current fail---"
                    End If
               
                   
                Tester.Print rv0, " \\CF :0 Unknow device, 1 pass ,2   fail, 3 over currnt, 4 preious slot fail"
                
               
                Tester.Print "LBA="; LBA
    
                 
                
AU6377ALFResult:
                        
                        If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                        ElseIf rv0 = 2 Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
                        ElseIf rv0 = 3 Then
                        
                             XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                            
                            
                        ElseIf rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
                
               
                  
               
                  
                CardResult = DO_WritePort(card, Channel_P1A, &H1)
                  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    
   End Sub

Public Function OpenDriver(Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

  
   
        Do
            DoEvents
            Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
            TmpString = GetDeviceName(Vid_PID)
        Loop While TmpString = "" And TimerCounter < 10
   
    '=======================================
    If OpenPipe = 0 Then
      OpenDriver = 2   ' Write fail
      Exit Function
    End If
     OpenDriver = 1
End Function
Public Function Read_Capacity(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte
Dim CBW(0 To 30) As Byte
Dim NumberOfBytesWritten As Long
Dim CBWDataTransferLen(0 To 3) As Byte
  
Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim i As Integer
Dim tmpV(0 To 2) As Long
Dim opcode As Byte

Dim CSW(0 To 12) As Byte

Dim NumberOfBytesRead As Long

Dim Capacity(0 To 7) As Byte

'Capacity(0) = &H0
'Capacity(1) = &H78
'Capacity(2) = &HFF
'Capacity(3) = &HFF
'Capacity(4) = &H0
'Capacity(5) = &H0
'Capacity(6) = &H2
'Capacity(7) = &H0

Capacity(0) = &H0
Capacity(1) = &H7C
Capacity(2) = &H7F
Capacity(3) = &HFD
Capacity(4) = &H0
Capacity(5) = &H0
Capacity(6) = &H2
Capacity(7) = &H0

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

 
CBW(8) = &H8  '00
CBW(9) = &H0  '08
CBW(10) = &H0 '00
CBW(11) = &H0 '00

'///////////////  CBW Flag
CBW(12) = &H80                 '80

'////////////// LUN
CBW(13) = Lun                    '00

'///////////// CBD Len
CBW(14) = &HA                '0a

'////////////  UFI command

CBW(15) = &H25
CBW(16) = Lun * 32
 
CBW(17) = &H0         '00
CBW(18) = &H0        '00
CBW(19) = &H0        '00
CBW(20) = &H0         '40

'/////////////  Reverve
CBW(21) = 0

'//////////// Transfer Len

 
CBW(22) = &H0     '00
CBW(23) = &H0     '04

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
 Read_Capacity = 0
 Exit Function
End If

'2. Readdata stage
 
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in




 
If result = 0 Then
  Read_Capacity = 0
 Exit Function
End If


For i = 0 To CBWDataTransferLength - 1
Debug.Print "k", i, Hex(ReadData(i)), Capacity(i)
'If ReadData(i) <> Capacity(i) Then
  
 ' Read_Capacity = 2  ' card format capacity has problem
  'Exit Function
'End If


Next i


'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
If result = 0 Then
 Read_Capacity = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
     Read_Capacity = 0
Else
      Read_Capacity = 1
   
End If

 
End Function

Public Function CBWTest_New_AU6390MB(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String, Flash As Byte) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long
Dim HalfFlashCapacity As Long
Dim OldLBa As Long

   CBWDataTransferLength = 1024 * 4  ' 8 sector, MB flash 2k/page, and it is 4 flash
                                    '  8 sector for 2 page for 2 flash
                                    ' the another 8 sector for another 2 flash set
   
'   For i = 0 To CBWDataTransferLength - 1
    
'         ReadData(i) = 0

'   Next

    If PreSlotStatus <> 1 Then
        CBWTest_New_AU6390MB = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New_AU6390MB = 0
    If Flash = 1 Or Flash = 0 Then
     If LBA > 3932160 Then
         LBA = 0
     End If
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
      CBWTest_New_AU6390MB = 0    ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_New_AU6390MB = 2    ' Write fail
      Exit Function
    End If
 
    '======================================
    
     ' for unitSpeed
    
    TmpInteger = TestUnitSpeed(Lun)
    
    If TmpInteger = 0 Then
        
       CBWTest_New_AU6390MB = 2    ' usb 2.0 high speed fail
       UsbSpeedTestResult = 2
       Exit Function
    End If
    
    
    
    TmpInteger = TestUnitReady(Lun)
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
           CBWTest_New_AU6390MB = 2   'Write fail
           Exit Function
        End If
        
    End If
    '======================================
   
    TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
      TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
   ' If TmpInteger = 0 Then
   '      CBWTest_New_AU6390MB = 2   'write fail
   '       Exit Function
   '  End If
    
      
   ' for AU6390MB to read capacity
   
   
     TmpInteger = Read_Capacity(LBA, Lun, 8)
      
    If TmpInteger = 0 Then
         CBWTest_New_AU6390MB = 3   'Read fail
          Exit Function
     ElseIf TmpInteger = 2 Then
           FlashCapacityError = 2
           CBWTest_New_AU6390MB = 3   'card format has problem
          Exit Function
          
     End If
      
      
     If Flash = 0 Then
      CBWTest_New_AU6390MB = 1
        Exit Function
     End If
      
 
 
    TmpInteger = Write_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_AU6390MB = 2   'write fail
        Exit Function
    End If
    
    TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_AU6390MB = 3     'Read fail
        Exit Function
    End If
     
    For i = 0 To CBWDataTransferLength - 1
    
        If ReadData(i) <> Pattern(i) Then
          CBWTest_New_AU6390MB = 3     'Read fail
          Exit Function
        End If
    
    Next
  
    ' another 2 flash R/W
  
   
    
    
    CBWTest_New_AU6390MB = 1
        
    
    End Function

 
   
Public Sub MultiSlotTestAU6378()
   
Dim TmpChip As String
Dim RomSelector As Byte
               
         
                
            '    PowerSet (1) ' for 3.3V , 2.5 V
 
  
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    
'If ChipName = "AU6368A" Then
'       CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 0111 1111
'           result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
'End If
 
        result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
    CardResult = DO_WritePort(card, Channel_P1A, &H3E)  ' 1111 1110
 
         CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 1111 1110
         CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' 1111 1110
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 1111 1110
 
  
   
 
  
  
  
  
  
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
              '  If TmpChip = "AU6377ALF25" Or TmpChip = "AU6378ALF22" Or TmpChip = "AU6378HLF22" Then
                  VidName = "vid_1984"
              '  Else
              '   VidName = "vid_058f"
               ' End If
                
               
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
                  TestResult = ""
                 
                    CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
                  
                   
               
                   
                  
                    
               
          
                   
                    
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
                 
                 
          
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 rv1 = CBWTest_NewAU6378AutoModeFail(1, rv0, VidName)    ' CF slot
            
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
                rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
                
                '============= SMC test begin =======================================
                If rv2 = 1 Then
                 '--- for SMC
                   CardResult = DO_WritePort(card, Channel_P1A, &H1C)  ' 0110 0100
                Call MsecDelay(0.2)
                
                CardResult = DO_WritePort(card, Channel_P1A, &H18)  ' 0110 0100
                Call MsecDelay(0.2)
                ClosePipe
                rv2 = CBWTest_New(2, rv2, VidName)
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
               
                 
                End If
                
                
           
                
               '=============== SMC test END ==================================================
               
               rv3 = CBWTest_New(3, rv2, VidName)
               ClosePipe
               Call LabelMenu(3, rv3, rv2)
           
       
                
             
                         ClosePipe
                         rv4 = CBWTest_New(4, rv3, VidName)
                          Call LabelMenu(10, rv4, rv3)
                          ClosePipe
          
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If rv4 = 1 And LightOff <> 252 Then
                            UsbSpeedTestResult = GPO_FAIL
                            rv4 = 2
                         End If
           
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                 Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                 Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                 Tester.Print rv3, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                 Tester.Print rv4, " \\MinSD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
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
                        ElseIf rv2 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv2 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                         ElseIf rv3 = WRITE_FAIL Or rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv3 = READ_FAIL Or rv4 = READ_FAIL Then
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

   Public Sub MultiSlotTestAU6378ALF24()
   
Dim TmpChip As String
Dim RomSelector As Byte
               
         
                
            '    PowerSet (1) ' for 3.3V , 2.5 V
 
  
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
'result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
'CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    
'If ChipName = "AU6368A" Then
'       CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 0111 1111
'           result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
'End If
 
        'result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
    'CardResult = DO_WritePort(card, Channel_P1A, &H3E)  ' 1111 1110
 
         'CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 1111 1110
         'CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' 1111 1110
         'result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.5)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 1111 1110
 
  
   
 
  
  
  
  
  
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
              '  If TmpChip = "AU6377ALF25" Or TmpChip = "AU6378ALF22" Or TmpChip = "AU6378HLF22" Then
                  VidName = "vid_1984"
              '  Else
              '   VidName = "vid_058f"
               ' End If
                
               
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
              
               rv4 = CBWTest_New_no_card(4, rv3, VidName)
                ' Tester.print "a4"
                ClosePipe
              Call LabelMenu(4, rv4, rv3)
                
 '================================= Test light off =============================
          
                
        
                
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
                 ElseIf rv3 = WRITE_FAIL Or rv4 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv3 = READ_FAIL Or rv4 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
                
                Tester.Print "Test Result"; TestResult
                       
       
                 
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv3, " \\MS :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv4, " \\MiniSD :0 Unknow device, 1 pass ,2 card change bit fail"
       
'====================================== Assing R/W test switch =====================================
                 
                If TestResult = "PASS" Then
                  TestResult = ""
                 
                    CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
                  
                   
               
                   
                  
                    
               
          
                   
                    
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
                 
                  If rv0 = 1 Then
                        rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                          
                          If rv0 <> 1 Then
                            Tester.Print "SD Bit width Fail"
                            rv0 = 2
                          End If
                          
               End If
                 
                 
                 
          
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 rv1 = CBWTest_NewAU6378AutoModeFail(1, rv0, VidName)    ' CF slot
            
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
                rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
                
                '============= SMC test begin =======================================
                If rv2 = 1 Then
                 '--- for SMC
                       CardResult = DO_WritePort(card, Channel_P1A, &H1C)  ' 0110 0100
                    Call MsecDelay(0.2)
                    
                    CardResult = DO_WritePort(card, Channel_P1A, &H18)  ' 0110 0100
                    Call MsecDelay(0.2)
                    ClosePipe
                    rv2 = CBWTest_New(2, rv2, VidName)
                    Call LabelMenu(2, rv2, rv1)
                    ClosePipe
               
                 
                End If
                
                
           
                
               '=============== SMC test END ==================================================
               
               rv3 = CBWTest_New(3, rv2, VidName)  ' MS slot
               
               If rv3 = 1 Then
                        rv3 = Read_MS_Speed(0, 0, 64, "4Bits")
                          
                          If rv3 <> 1 Then
                            Tester.Print "MS Bit width Fail"
                            rv3 = 2
                          End If
                          
               End If
               
               ClosePipe
               Call LabelMenu(3, rv3, rv2)
           
       
                
             
                         ClosePipe
                         rv4 = CBWTest_New(4, rv3, VidName)
                         
                       
                          Call LabelMenu(4, rv4, rv3)
                          ClosePipe
                          
                          
                         
                          
          
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If rv4 = 1 And LightOff <> 252 Then
                            UsbSpeedTestResult = GPO_FAIL
                            rv4 = 2
                         End If
           
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                 Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                 Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                 Tester.Print rv3, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                 Tester.Print rv4, " \\MinSD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
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
                        ElseIf rv2 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv2 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                         ElseIf rv3 = WRITE_FAIL Or rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv3 = READ_FAIL Or rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
                
               
                  
                End If
                CardResult = DO_WritePort(card, Channel_P1A, &H80)
                  'result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
                    'CardResult = DO_WritePort(card, Channel_P1B, &H0)
   End Sub

Public Sub AU6378ALS14TestSub()
   
Dim TmpChip As String
Dim RomSelector As Byte
Dim VidName As String
                
         
                
            '    PowerSet (1) ' for 3.3V , 2.5 V
 
  
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
Routine_Label:

ReaderExist = 0


Call PowerSet2(1, "3.3", "0.6", 1, "3.3", "0.6", 1)
Tester.Print "AU6378AL ST1: 3.3V Begin Test ..."
           

CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 1111 1110
  
'======================== Begin test ============================================
                  
VidName = "vid_1984"
Call MsecDelay(0.4)
WaitDevOn (VidName)
Call MsecDelay(0.1)
LBA = LBA + 1
                
                
'//////////////////////////////////////////////////
'
'   no card insert
'
'/////////////////////////////////////////////////

Tester.Print LBA

rv0 = 0
rv1 = 0
rv2 = 0
rv3 = 0
rv4 = 0
               
ClosePipe
rv0 = CBWTest_New_no_card(0, 1, VidName)
Call LabelMenu(0, rv0, 1)
ClosePipe
rv1 = CBWTest_New_no_card(1, rv0, VidName)
Call LabelMenu(1, rv1, rv0)
ClosePipe

rv2 = CBWTest_New_no_card(2, rv1, VidName)
Call LabelMenu(2, rv2, rv1)
ClosePipe
rv3 = CBWTest_New_no_card(3, rv2, VidName)
ClosePipe
Call LabelMenu(3, rv3, rv2)

rv4 = CBWTest_New_no_card(4, rv3, VidName)
ClosePipe
Call LabelMenu(4, rv4, rv3)

'================================= Test light off =============================
   
                
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
ElseIf rv3 = WRITE_FAIL Or rv4 = WRITE_FAIL Then
    MSWriteFail = MSWriteFail + 1
    TestResult = "MS_WF"
ElseIf rv3 = READ_FAIL Or rv4 = READ_FAIL Then
    MSReadFail = MSReadFail + 1
    TestResult = "MS_RF"
ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
    TestResult = "PASS"
Else
    TestResult = "Bin2"
End If

Tester.Print "Test Result"; TestResult
      
Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 card change bit fail"
Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 card change bit fail"
Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 card change bit fail"
Tester.Print rv3, " \\MS :0 Unknow device, 1 pass ,2 card change bit fail"
Tester.Print rv4, " \\MiniSD :0 Unknow device, 1 pass ,2 card change bit fail"
       
'====================================== Assing R/W test switch =====================================
                 
If TestResult = "PASS" Then
    TestResult = ""
                 
    CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
                   
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
    rv4 = 0
    
    Tester.Label3.BackColor = RGB(255, 255, 255)
    Tester.Label4.BackColor = RGB(255, 255, 255)
    Tester.Label5.BackColor = RGB(255, 255, 255)
    Tester.Label6.BackColor = RGB(255, 255, 255)
    Tester.Label7.BackColor = RGB(255, 255, 255)
    Tester.Label8.BackColor = RGB(255, 255, 255)
                
    ClosePipe
    rv0 = CBWTest_New(0, 1, VidName)     ' SD slot
    
    If rv0 = 1 Then
        rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
        
        If rv0 <> 1 Then
            Tester.Print "SD Bit width Fail"
            rv0 = 2
        End If
    End If
          
    Call LabelMenu(0, rv0, 1)
    ClosePipe
    rv1 = CBWTest_NewAU6378AutoModeFail(1, rv0, VidName)    ' CF slot
    
    Call LabelMenu(1, rv1, rv0)
    ClosePipe
    
    rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
    Call LabelMenu(2, rv2, rv1)
    ClosePipe
                
    '============= SMC test begin =======================================
    If rv2 = 1 Then
    '--- for SMC
        CardResult = DO_WritePort(card, Channel_P1A, &H1C)  ' 0110 0100
        Call MsecDelay(0.2)
        
        CardResult = DO_WritePort(card, Channel_P1A, &H18)  ' 0110 0100
        Call MsecDelay(0.2)
        ClosePipe
        rv2 = CBWTest_New(2, rv2, VidName)
        Call LabelMenu(2, rv2, rv1)
        ClosePipe
    End If
                
                
    '=============== SMC test END ==================================================
    
    rv3 = CBWTest_New(3, rv2, VidName)  ' MS slot
    
    If rv3 = 1 Then
        rv3 = Read_MS_Speed(0, 0, 64, "4Bits")
        
        If rv3 <> 1 Then
            Tester.Print "MS Bit width Fail"
            rv3 = 2
        End If
    End If
               
    ClosePipe
    Call LabelMenu(3, rv3, rv2)
             
    ClosePipe
    rv4 = CBWTest_New(4, rv3, VidName)
    
    Call LabelMenu(4, rv4, rv3)
    ClosePipe
          
    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
    If rv4 = 1 And LightOff <> 252 Then
        UsbSpeedTestResult = GPO_FAIL
        rv4 = 2
    End If
           
                
    Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    Tester.Print rv3, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    Tester.Print rv4, " \\MinSD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

End If
                 
AU6377ALFResult:

    
    Call PowerSet2(0, "0.0", "1.0", 1, "0.0", "1.0", 1)
    CardResult = DO_WritePort(card, Channel_P1A, &H80)
               
    WaitDevOFF (VidName)
    WaitDevOFF ("vid_0a48")
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
     ElseIf rv3 = WRITE_FAIL Or rv4 = WRITE_FAIL Then
        MSWriteFail = MSWriteFail + 1
        TestResult = "MS_WF"
    ElseIf rv3 = READ_FAIL Or rv4 = READ_FAIL Then
        MSReadFail = MSReadFail + 1
        TestResult = "MS_RF"
    ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
         TestResult = "PASS"
    Else
        TestResult = "Bin2"
      
    End If
                                  
        
End Sub

Public Sub AU6378ALF04TestSub()
   
Dim TmpChip As String
Dim RomSelector As Byte
Dim HV_Done_Flag As Boolean
Dim HV_Result As String
Dim LV_Result As String
Dim VidName As String
                
         
                
            '    PowerSet (1) ' for 3.3V , 2.5 V
 
  
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
Routine_Label:

ReaderExist = 0

'Call MsecDelay(0.5)



Call MsecDelay(0.3)
If Not HV_Done_Flag Then
    Call PowerSet2(1, "3.6", "0.6", 1, "3.6", "0.6", 1)
    'Call MsecDelay(0.2)
    Tester.Print "AU6378AL : HV(3.6) Begin Test ..."
Else
    Call PowerSet2(1, "3.3", "0.6", 1, "3.3", "0.6", 1)
    'Call MsecDelay(0.2)
    Tester.Print vbCrLf & "AU6378AL : LV(3.3) Begin Test ..."
End If
           

CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 1111 1110
  
'======================== Begin test ============================================
                  
VidName = "vid_1984"
Call MsecDelay(0.4)
WaitDevOn (VidName)
Call MsecDelay(0.1)
LBA = LBA + 1
                
                
'//////////////////////////////////////////////////
'
'   no card insert
'
'/////////////////////////////////////////////////

Tester.Print LBA

rv0 = 0
rv1 = 0
rv2 = 0
rv3 = 0
rv4 = 0
               
ClosePipe
rv0 = CBWTest_New_no_card(0, 1, VidName)
Call LabelMenu(0, rv0, 1)
ClosePipe
rv1 = CBWTest_New_no_card(1, rv0, VidName)
Call LabelMenu(1, rv1, rv0)
ClosePipe

rv2 = CBWTest_New_no_card(2, rv1, VidName)
Call LabelMenu(2, rv2, rv1)
ClosePipe
rv3 = CBWTest_New_no_card(3, rv2, VidName)
ClosePipe
Call LabelMenu(3, rv3, rv2)

rv4 = CBWTest_New_no_card(4, rv3, VidName)
ClosePipe
Call LabelMenu(4, rv4, rv3)

'================================= Test light off =============================
   
                
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
ElseIf rv3 = WRITE_FAIL Or rv4 = WRITE_FAIL Then
    MSWriteFail = MSWriteFail + 1
    TestResult = "MS_WF"
ElseIf rv3 = READ_FAIL Or rv4 = READ_FAIL Then
    MSReadFail = MSReadFail + 1
    TestResult = "MS_RF"
ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
    TestResult = "PASS"
Else
    TestResult = "Bin2"
End If

Tester.Print "Test Result"; TestResult
      
Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 card change bit fail"
Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 card change bit fail"
Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 card change bit fail"
Tester.Print rv3, " \\MS :0 Unknow device, 1 pass ,2 card change bit fail"
Tester.Print rv4, " \\MiniSD :0 Unknow device, 1 pass ,2 card change bit fail"
       
'====================================== Assing R/W test switch =====================================
                 
If TestResult = "PASS" Then
    TestResult = ""
                 
    CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
                   
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
    rv4 = 0
    
    Tester.Label3.BackColor = RGB(255, 255, 255)
    Tester.Label4.BackColor = RGB(255, 255, 255)
    Tester.Label5.BackColor = RGB(255, 255, 255)
    Tester.Label6.BackColor = RGB(255, 255, 255)
    Tester.Label7.BackColor = RGB(255, 255, 255)
    Tester.Label8.BackColor = RGB(255, 255, 255)
                
    ClosePipe
    rv0 = CBWTest_New(0, 1, VidName)     ' SD slot
    
    If rv0 = 1 Then
        rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
        
        If rv0 <> 1 Then
            Tester.Print "SD Bit width Fail"
            rv0 = 2
        End If
    End If
          
    Call LabelMenu(0, rv0, 1)
    ClosePipe
    rv1 = CBWTest_NewAU6378AutoModeFail(1, rv0, VidName)    ' CF slot
    
    Call LabelMenu(1, rv1, rv0)
    ClosePipe
    
    rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
    Call LabelMenu(2, rv2, rv1)
    ClosePipe
                
    '============= SMC test begin =======================================
    If rv2 = 1 Then
    '--- for SMC
        CardResult = DO_WritePort(card, Channel_P1A, &H1C)  ' 0110 0100
        Call MsecDelay(0.2)
        
        CardResult = DO_WritePort(card, Channel_P1A, &H18)  ' 0110 0100
        Call MsecDelay(0.2)
        ClosePipe
        rv2 = CBWTest_New(2, rv2, VidName)
        Call LabelMenu(2, rv2, rv1)
        ClosePipe
    End If
                
                
    '=============== SMC test END ==================================================
    
    rv3 = CBWTest_New(3, rv2, VidName)  ' MS slot
    
    If rv3 = 1 Then
        rv3 = Read_MS_Speed(0, 0, 64, "4Bits")
        
        If rv3 <> 1 Then
            Tester.Print "MS Bit width Fail"
            rv3 = 2
        End If
    End If
               
    ClosePipe
    Call LabelMenu(3, rv3, rv2)
             
    ClosePipe
    rv4 = CBWTest_New(4, rv3, VidName)
    
    Call LabelMenu(4, rv4, rv3)
    ClosePipe
          
    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
    If rv4 = 1 And LightOff <> 252 Then
        UsbSpeedTestResult = GPO_FAIL
        rv4 = 2
    End If
           
                
    Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    Tester.Print rv3, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    Tester.Print rv4, " \\MinSD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

End If
                 
AU6377ALFResult:

    
    Call PowerSet2(0, "0.0", "1.0", 1, "0.0", "1.0", 1)
    CardResult = DO_WritePort(card, Channel_P1A, &H80)
               
    WaitDevOFF (VidName)
    WaitDevOFF ("vid_0a48")
    Call MsecDelay(0.3)
                
    If HV_Done_Flag = False Then
        If rv0 <> 1 Then
            HV_Result = "Bin2"
            Tester.Print "HV Unknow"
        ElseIf rv0 * rv1 * rv2 * rv3 * rv4 <> 1 Then
            HV_Result = "Fail"
            Tester.Print "HV Fail"
        ElseIf rv0 * rv1 * rv2 * rv3 * rv4 = 1 Then
            HV_Result = "PASS"
            Tester.Print "HV PASS"
        End If
        
        HV_Done_Flag = True
        GoTo Routine_Label
    Else
        If rv0 <> 1 Then
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
   
   Public Sub MultiSlotTestAU6378RLF24TestSub()
   
Dim TmpChip As String
Dim RomSelector As Byte
               
         
                
            '    PowerSet (1) ' for 3.3V , 2.5 V
 
  
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    
'If ChipName = "AU6368A" Then
'       CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 0111 1111
'           result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
'End If
 
        result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
    CardResult = DO_WritePort(card, Channel_P1A, &H3E)  ' 1111 1110
 
         CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 1111 1110
         CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' 1111 1110
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 1111 1110
 
  
   
 
  
  
  
  
  
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
              '  If TmpChip = "AU6377ALF25" Or TmpChip = "AU6378ALF22" Or TmpChip = "AU6378HLF22" Then
                  VidName = "vid_1984"
              '  Else
              '   VidName = "vid_058f"
               ' End If
                
               
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
                 rv4 = CBWTest_New_no_card(4, rv3, VidName)
                ' Tester.print "a4"
                ClosePipe
              Call LabelMenu(4, rv4, rv3)
                
 '================================= Test light off =============================
          
                
        
                
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
                 ElseIf rv3 = WRITE_FAIL Or rv4 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv3 = READ_FAIL Or rv4 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
                
                Tester.Print "Test Result"; TestResult
                       
       
                 
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv3, " \\MS :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv4, " \\Mini SD :0 Unknow device, 1 pass ,2 card change bit fail"
'====================================== Assing R/W test switch =====================================
                 
                If TestResult = "PASS" Then
                  TestResult = ""
                 
                    CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
                  
                   
               
                   
                  
                    
               
          
                   
                    
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
                rv4 = 0
                Tester.Label3.BackColor = RGB(255, 255, 255)
                Tester.Label4.BackColor = RGB(255, 255, 255)
                Tester.Label5.BackColor = RGB(255, 255, 255)
                Tester.Label6.BackColor = RGB(255, 255, 255)
                Tester.Label7.BackColor = RGB(255, 255, 255)
                Tester.Label8.BackColor = RGB(255, 255, 255)
                
                
                
                ClosePipe
                 rv0 = CBWTest_New(0, 1, VidName)     ' SD slot
                 
                    If rv0 = 1 Then
                        rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                          
                          If rv0 <> 1 Then
                            Tester.Print "SD Bit width Fail"
                            rv0 = 2
                          End If
                          
                End If
                 
                 
          
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 rv1 = CBWTest_NewAU6378AutoModeFail(1, rv0, VidName)    ' CF slot
            
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
                rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
                
                '============= SMC test begin =======================================
                If rv2 = 1 Then
                 '--- for SMC
                   CardResult = DO_WritePort(card, Channel_P1A, &H1C)  ' 0110 0100
                Call MsecDelay(0.2)
                
                CardResult = DO_WritePort(card, Channel_P1A, &H18)  ' 0110 0100
                Call MsecDelay(0.2)
                ClosePipe
                rv2 = CBWTest_New(2, rv2, VidName)
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
               
                 
                End If
                
                
           
                
               '=============== SMC test END ==================================================
               
               rv3 = CBWTest_New(3, rv2, VidName)
                 If rv3 = 1 Then
                        rv3 = Read_MS_Speed(0, 0, 64, "4Bits")
                          
                          If rv3 <> 1 Then
                            Tester.Print "MS Bit width Fail"
                            rv3 = 2
                          End If
                          
               End If
               
               
               ClosePipe
               Call LabelMenu(3, rv3, rv2)
           
       
                
             
                         ClosePipe
                         rv4 = CBWTest_New(4, rv3, VidName)
                          Call LabelMenu(4, rv4, rv3)
                          ClosePipe
          
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If rv4 = 1 And LightOff <> 254 Then
                            UsbSpeedTestResult = GPO_FAIL
                            rv4 = 2
                         End If
           
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                 Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                 Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                 Tester.Print rv3, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                 Tester.Print rv4, " \\MinSD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
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
                        ElseIf rv2 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv2 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                         ElseIf rv3 = WRITE_FAIL Or rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv3 = READ_FAIL Or rv4 = READ_FAIL Then
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
  
   
   
    Public Sub MultiSlotTestAU6378RLTestSub()
   
Dim TmpChip As String
Dim RomSelector As Byte
               
         
                
            '    PowerSet (1) ' for 3.3V , 2.5 V
 
  
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    
'If ChipName = "AU6368A" Then
'       CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 0111 1111
'           result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
'End If
 
        result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
    CardResult = DO_WritePort(card, Channel_P1A, &H3E)  ' 1111 1110
 
         CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 1111 1110
         CardResult = DO_WritePort(card, Channel_P1A, &HFF)  ' 1111 1110
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.3)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 1111 1110
 
  
   
 
  
  
  
  
  
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
              '  If TmpChip = "AU6377ALF25" Or TmpChip = "AU6378ALF22" Or TmpChip = "AU6378HLF22" Then
                  VidName = "vid_1984"
              '  Else
              '   VidName = "vid_058f"
               ' End If
                
               
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
                  TestResult = ""
                 
                    CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
                  
                   
               
                   
                  
                    
               
          
                   
                    
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
                 
                 
          
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 rv1 = CBWTest_NewAU6378AutoModeFail(1, rv0, VidName)    ' CF slot
            
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
              
                rv2 = CBWTest_New(2, rv1, VidName)    'XD slot
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
                
                '============= SMC test begin =======================================
                If rv2 = 1 Then
                 '--- for SMC
                   CardResult = DO_WritePort(card, Channel_P1A, &H1C)  ' 0110 0100
                Call MsecDelay(0.2)
                
                CardResult = DO_WritePort(card, Channel_P1A, &H18)  ' 0110 0100
                Call MsecDelay(0.2)
                ClosePipe
                rv2 = CBWTest_New(2, rv2, VidName)
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
               
                 
                End If
                
                
           
                
               '=============== SMC test END ==================================================
               
               rv3 = CBWTest_New(3, rv2, VidName)
               ClosePipe
               Call LabelMenu(3, rv3, rv2)
           
       
                
             
                         ClosePipe
                         rv4 = CBWTest_New(4, rv3, VidName)
                          Call LabelMenu(10, rv4, rv3)
                          ClosePipe
          
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If rv4 = 1 And LightOff <> 254 Then
                            UsbSpeedTestResult = GPO_FAIL
                            rv4 = 2
                         End If
           
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                 Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                 Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                 Tester.Print rv3, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                 Tester.Print rv4, " \\MinSD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
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
                        ElseIf rv2 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv2 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                         ElseIf rv3 = WRITE_FAIL Or rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv3 = READ_FAIL Or rv4 = READ_FAIL Then
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


Public Function CBWTest_New_ALPS(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

   CBWDataTransferLength = 1024
 
'   For i = 0 To CBWDataTransferLength - 1
    
'         ReadData(i) = 0

'   Next

    If PreSlotStatus <> 1 Then
        CBWTest_New_ALPS = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New_ALPS = 0
  '   If Lba > 25 * 1024 Then
  '       Lba = 0
  '   End If
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
      CBWTest_New_ALPS = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_New_ALPS = 2   ' Write fail
      Exit Function
    End If
    '======================================
    
      TmpInteger = TestUnitSpeed(Lun)
    
    If TmpInteger = 0 Then
        
       CBWTest_New_ALPS = 2   ' usb 2.0 high speed fail
       UsbSpeedTestResult = 2
       Exit Function
    End If
    TmpInteger = 0
    
    
    
    TmpInteger = TestUnitReady(Lun)
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
           CBWTest_New_ALPS = 2  'Write fail
           Exit Function
        End If
        
    End If
    '======================================
   
    TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
      
    If TmpInteger = 0 Then
        CBWTest_New_ALPS = 2  'write fail
        Exit Function
    End If
    
      
    TmpInteger = Write_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_ALPS = 2  'write fail
        Exit Function
    End If
    
    
    ' For ALPS  Issue
    TmpInteger = Read_Data(LBA + CLng(80000), Lun, CBWDataTransferLength)
    
    
    TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_ALPS = 3    'Read fail
        Exit Function
    End If
     
    For i = 0 To CBWDataTransferLength - 1
    
        If ReadData(i) <> Pattern(i) Then
          CBWTest_New_ALPS = 3    'Read fail
          Exit Function
        End If
    
    Next
    
    CBWTest_New_ALPS = 1
        
    
    End Function


Public Sub MultiSlotTestAU6378FLF24Test()
   
Dim TmpChip As String
Dim RomSelector As Byte
 
     TmpChip = ChipName
     ChipName = "AU6376"
 
            
                'GPIO control setting
 
                
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
CardResult = DO_WritePort(card, Channel_P1B, &H0)
 
 
        result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
  
 
                  
  
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 1111 1110
         CardResult = DO_WritePort(card, Channel_P1A, &HFE)  ' 1111 1110
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.1)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 1111 1110
  
  
  
 
  
  
  
  
  
'======================== Begin test ============================================
                  
                Call MsecDelay(0.8)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
                Tester.Print LBA
                 
                  VidName = "vid_1984"
                
              
                ClosePipe
                 rv0 = CBWTest_New_no_card(0, 1, VidName)
                'Tester.print "a1"
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                rv1 = CBWTest_New_no_card(4, rv0, VidName)  ' T-flash
               '  Tester.print "a2"
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
                
                
                
 '================================= Test light off =============================
                
               
                
                
              
                ' test chip
                      '    CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
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
                ElseIf rv1 * rv0 = PASS Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
                
                Tester.Print "Test Result"; TestResult
                       
       
                 
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv1, " \\T-Flash :0 Unknow device, 1 pass ,2 card change bit fail"
                
                 
'====================================== Assing R/W test switch =====================================
                 
                If TestResult = "PASS" Then
           
                    TestResult = ""
             
                    CardResult = DO_WritePort(card, Channel_P1A, &H7E)  ' 0110 0100
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
                 
                 
                   If rv0 = 1 Then
                        rv0 = Read_SD_Speed(0, 0, 64, "8Bits")
                          
                          If rv0 <> 1 Then
                            Tester.Print "SD Bit width Fail"
                            rv0 = 2
                          End If
                          
                End If
                 
                Call LabelMenu(0, rv0, 1)
                
               
             
                 ClosePipe
                 rv1 = CBWTest_New(4, rv0, VidName)
                
                 ClosePipe
          
               
                      
                 Call LabelMenu(1, rv1, rv0)
               
                 If rv1 = 2 Or rv1 = 3 Then
                 Tester.Label9.Caption = "T-Flash fail"
                 End If
                 
                 
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                 If rv1 = 1 And LightOn <> 254 Then
                      UsbSpeedTestResult = GPO_FAIL
                       Tester.Label9.Caption = "GPO_FAIL fail"
                      rv1 = 2
                  End If
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\T-Flash :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                
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
                        ElseIf rv2 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv2 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                         ElseIf rv3 = WRITE_FAIL Or rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv3 = READ_FAIL Or rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
                
               
                  
                End If
                CardResult = DO_WritePort(card, Channel_P1A, &H1)
                  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
   End Sub
Public Sub MultiSlotTestAU6378FLTest()
   
Dim TmpChip As String
Dim RomSelector As Byte
 
     TmpChip = ChipName
     ChipName = "AU6376"
 
            
                'GPIO control setting
 
                
'==================================== Switch assign ==========================================
            
            
If PCI7248InitFinish = 0 Then
      PCI7248Exist
End If
                
result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
CardResult = DO_WritePort(card, Channel_P1B, &H0)
 
 
        result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
  
 
                  
  
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 1111 1110
         CardResult = DO_WritePort(card, Channel_P1A, &HFE)  ' 1111 1110
         result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
         Call MsecDelay(0.1)
                    
         CardResult = DO_WritePort(card, Channel_P1A, &H7F)  ' 1111 1110
  
  
  
 
  
  
  
  
  
'======================== Begin test ============================================
                  
                Call MsecDelay(0.8)
               
               LBA = LBA + 1
                
                
                '//////////////////////////////////////////////////
                '
                '   no card insert
                '
                '/////////////////////////////////////////////////
                
                Dim VidName As String
                Tester.Print LBA
                 
                  VidName = "vid_1984"
                
              
                ClosePipe
                 rv0 = CBWTest_New_no_card(0, 1, VidName)
                'Tester.print "a1"
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                rv1 = CBWTest_New_no_card(4, rv0, VidName)  ' T-flash
               '  Tester.print "a2"
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
                
                
                
 '================================= Test light off =============================
                
               
                
                
              
                ' test chip
                      '    CardResult = DO_WritePort(card, Channel_P1A, &H14)  ' 0110 0100
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
                ElseIf rv1 * rv0 = PASS Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
                
                Tester.Print "Test Result"; TestResult
                       
       
                 
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 card change bit fail"
                Tester.Print rv1, " \\T-Flash :0 Unknow device, 1 pass ,2 card change bit fail"
                
                 
'====================================== Assing R/W test switch =====================================
                 
                If TestResult = "PASS" Then
           
                    TestResult = ""
             
                    CardResult = DO_WritePort(card, Channel_P1A, &H7E)  ' 0110 0100
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
                 
                Call LabelMenu(0, rv0, 1)
                
               
             
                 ClosePipe
                 rv1 = CBWTest_New(4, rv0, VidName)
                
                 ClosePipe
          
               
                      
                 Call LabelMenu(1, rv1, rv0)
               
                 If rv1 = 2 Or rv1 = 3 Then
                 Tester.Label9.Caption = "T-Flash fail"
                 End If
                 
                 
                   CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                 If rv1 = 1 And LightOn <> 254 Then
                      UsbSpeedTestResult = GPO_FAIL
                       Tester.Label9.Caption = "GPO_FAIL fail"
                      rv1 = 2
                  End If
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\T-Flash :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                
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
                        ElseIf rv2 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv2 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                         ElseIf rv3 = WRITE_FAIL Or rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv3 = READ_FAIL Or rv4 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                        ElseIf rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
                
               
                  
                End If
                CardResult = DO_WritePort(card, Channel_P1A, &H1)
                  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
   End Sub
   
Public Sub MultiSlotTest()
   
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
 If ChipName = "AU6370DLF20" Or ChipName = "AU6378ALF20" Then
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
                  
 
                  
 If TmpChip = "AU6378ALF20" Then
                    
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
                If TmpChip = "AU6377ALF25" Then
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
                
                
                If TmpChip = "AU6378ALF20" Then
                
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
                   '
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
                   
                    If TmpChip = "AU6378ALF20" Then
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
                 
                 If rv0 = 1 And Left(TmpChip, 10) = "AU6375HLF2" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                    
                    ' AU6375 ram unstable
                    
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 5
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6375(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                    
                    
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
               
                If rv2 = 1 And TmpChip = "AU6378ALF20" Then         '--- for SMC
                
                CardResult = DO_WritePort(card, Channel_P1A, &H18)  ' 0110 0100
                Call MsecDelay(0.5)
                ClosePipe
                rv2 = CBWTest_New(2, rv2, VidName)
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
                 End If
                 
              If rv2 = 1 And Left(TmpChip, 10) = "AU6377ALF2" And TmpChip <> "AU6377ALF21" Then           '--- for SMC
                
                CardResult = DO_WritePort(card, Channel_P1A, &H8 + RomSelector) ' 0110 0100
                Call MsecDelay(0.5)
                ClosePipe
                rv2 = CBWTest_New(2, rv2, VidName)
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
              End If
                
                
               If rv2 = 1 And (TmpChip = "AU6376JLF20") Then      '--- for SMC
                
                  CardResult = DO_WritePort(card, Channel_P1A, &H18)   ' 0110 0100
                  Call MsecDelay(0.5)
                  CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                  Call MsecDelay(0.5)
                 ClosePipe
                 rv2 = CBWTest_New(2, rv2, VidName)
                 Call LabelMenu(2, rv2, rv1)
                 ClosePipe
               End If
               
               
                
                
               '=============== SMC test END ==================================================
               
               rv3 = CBWTest_New(3, rv2, VidName)
               ClosePipe
               Call LabelMenu(3, rv3, rv2)
           
                 
               If TmpChip = "AU6375HLF21" Then
               
                 If rv0 = 1 Then
                   
                    ClosePipe
                     rv0 = CBWTest_New_AU6375IncPattern(0, 1, VidName)
                     Call LabelMenu(0, rv0, 1)
                     ClosePipe
                 End If
                
                End If
                
                 
                 
                If Left(TmpChip, 10) = "AU6377ALF2" Then
                
                ' test chip
                         ClosePipe
                         rv4 = CBWTest_New(4, rv3, VidName)   'MMC test
                          Call LabelMenu(10, rv4, rv3)
                          ClosePipe
          
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If LightOff <> 127 Then
                            UsbSpeedTestResult = GPO_FAIL
                            rv4 = 2
                         End If
          
                End If
                 
                If TmpChip = "AU6378ALF20" Then
                
             
                         ClosePipe
                         rv4 = CBWTest_New(4, rv3, VidName)
                          Call LabelMenu(10, rv4, rv3)
                          ClosePipe
          
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If rv4 = 1 And LightOff <> 252 Then
                            UsbSpeedTestResult = GPO_FAIL
                            rv4 = 2
                         End If
          
                End If
                 
                 
                    
                  If ChipName = "AU6376" And TmpChip = "AU6370DLF20" Then
                  Call MsecDelay(0.1)
                  rv4 = 1
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 If LightOff <> 254 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                 End If
                     Call LabelMenu(3, rv4, rv3)
                     
                        
                 End If
                 
                 
                 If ChipName = "AU6368A1" Then
                 Call MsecDelay(0.1)
                  rv4 = 1
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 If LightOff <> 192 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                 End If
                     Call LabelMenu(3, rv4, rv3)
                     
                       CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 
                 End If
                 
                  If ChipName = "AU6376" And (Left(TmpChip, 10) <> "AU6377ALF2" And TmpChip <> "AU6378ALF20") Then
                  Call MsecDelay(0.1)
                  rv4 = 1
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                               If TmpChip = "AU6370DLF20" Or TmpChip = "AU6376JLF20" Then
                               
                                  If LightOff <> 254 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                 End If
                               Else
                                 If LightOff <> 252 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                 End If
                              End If
                     Call LabelMenu(3, rv4, rv3)
                     
                        
                 End If
                 
                If ChipName = "AU6368A" Then
                    If rv3 = 1 Then
                           CardResult = DO_WritePort(card, Channel_P1A, &H74)  ' 0111 0100
                           Call MsecDelay(0.1)
                           CardResult = DO_WritePort(card, Channel_P1A, &H54)  ' 0101 0100
                           Call MsecDelay(0.1)
                           rv4 = CBWTest_New(3, rv3, VidName)
                             ClosePipe
                           If rv4 = 1 Then
                                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 If LightOff <> 132 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                    rv4 = 2
                                 End If
                             End If
                         Else
                         rv4 = 4
                         End If
                         Call LabelMenu(3, rv4, rv3)
                 End If
                
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
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
                        ElseIf rv2 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv2 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                         ElseIf rv3 = WRITE_FAIL Or rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv3 = READ_FAIL Or rv4 = READ_FAIL Then
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
   
Public Sub MultiSlotTestAU6376ALF20()
   
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
 If ChipName = "AU6370DLF20" Or ChipName = "AU6378ALF20" Then
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
                  
 
                  
 If TmpChip = "AU6378ALF20" Then
                    
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
                If TmpChip = "AU6377ALF25" Then
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
                
                
                If TmpChip = "AU6378ALF20" Then
                
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
                   '
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
                   
                    If TmpChip = "AU6378ALF20" Then
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
                 
                 If rv0 = 1 And Left(TmpChip, 10) = "AU6375HLF2" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                    
                    ' AU6375 ram unstable
                    
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 5
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6375(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                    
                    
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
               
                If rv2 = 1 And TmpChip = "AU6378ALF20" Then         '--- for SMC
                
                CardResult = DO_WritePort(card, Channel_P1A, &H18)  ' 0110 0100
                Call MsecDelay(0.5)
                ClosePipe
                rv2 = CBWTest_New(2, rv2, VidName)
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
                 End If
                 
              If rv2 = 1 And Left(TmpChip, 10) = "AU6377ALF2" And TmpChip <> "AU6377ALF21" Then           '--- for SMC
                
                CardResult = DO_WritePort(card, Channel_P1A, &H8 + RomSelector) ' 0110 0100
                Call MsecDelay(0.5)
                ClosePipe
                rv2 = CBWTest_New(2, rv2, VidName)
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
              End If
                
                
               If rv2 = 1 And (TmpChip = "AU6376JLF20") Then      '--- for SMC
                
                  CardResult = DO_WritePort(card, Channel_P1A, &H18)   ' 0110 0100
                  Call MsecDelay(0.5)
                  CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                  Call MsecDelay(0.5)
                 ClosePipe
                 rv2 = CBWTest_New(2, rv2, VidName)
                 Call LabelMenu(2, rv2, rv1)
                 ClosePipe
               End If
               
               
                CardResult = DO_WritePort(card, Channel_P1A, &H18)   ' 0110 0100
                  Call MsecDelay(0.5)
                  CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                  Call MsecDelay(0.5)
                 ClosePipe
                 rv2 = CBWTest_New(2, rv2, VidName)
                 Call LabelMenu(2, rv2, rv1)
                 ClosePipe
               
          
               '=============== SMC test END ==================================================
               
               rv3 = CBWTest_New(3, rv2, VidName)
               ClosePipe
               Call LabelMenu(3, rv3, rv2)
           
                 
               If TmpChip = "AU6375HLF21" Then
               
                 If rv0 = 1 Then
                   
                    ClosePipe
                     rv0 = CBWTest_New_AU6375IncPattern(0, 1, VidName)
                     Call LabelMenu(0, rv0, 1)
                     ClosePipe
                 End If
                
                End If
                
                 
                 
                If Left(TmpChip, 10) = "AU6377ALF2" Then
                
                ' test chip
                         ClosePipe
                         rv4 = CBWTest_New(4, rv3, VidName)   'MMC test
                          Call LabelMenu(10, rv4, rv3)
                          ClosePipe
          
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If LightOff <> 127 Then
                            UsbSpeedTestResult = GPO_FAIL
                            rv4 = 2
                         End If
          
                End If
                 
                If TmpChip = "AU6378ALF20" Then
                
             
                         ClosePipe
                         rv4 = CBWTest_New(4, rv3, VidName)
                          Call LabelMenu(10, rv4, rv3)
                          ClosePipe
          
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If rv4 = 1 And LightOff <> 252 Then
                            UsbSpeedTestResult = GPO_FAIL
                            rv4 = 2
                         End If
          
                End If
                 
                 
                    
                  If ChipName = "AU6376" And TmpChip = "AU6370DLF20" Then
                  Call MsecDelay(0.1)
                  rv4 = 1
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 If LightOff <> 254 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                 End If
                     Call LabelMenu(3, rv4, rv3)
                     
                        
                 End If
                 
                 
                 If ChipName = "AU6368A1" Then
                 Call MsecDelay(0.1)
                  rv4 = 1
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 If LightOff <> 192 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                 End If
                     Call LabelMenu(3, rv4, rv3)
                     
                       CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 
                 End If
                 
                  If ChipName = "AU6376" And (Left(TmpChip, 10) <> "AU6377ALF2" And TmpChip <> "AU6378ALF20") Then
                  Call MsecDelay(0.1)
                  rv4 = 1
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                               If TmpChip = "AU6370DLF20" Or TmpChip = "AU6376JLF20" Then
                               
                                  If LightOff <> 254 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                 End If
                               Else
                                 If LightOff <> 252 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                 End If
                              End If
                     Call LabelMenu(3, rv4, rv3)
                     
                        
                 End If
                 
                If ChipName = "AU6368A" Then
                    If rv3 = 1 Then
                           CardResult = DO_WritePort(card, Channel_P1A, &H74)  ' 0111 0100
                           Call MsecDelay(0.1)
                           CardResult = DO_WritePort(card, Channel_P1A, &H54)  ' 0101 0100
                           Call MsecDelay(0.1)
                           rv4 = CBWTest_New(3, rv3, VidName)
                             ClosePipe
                           If rv4 = 1 Then
                                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 If LightOff <> 132 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                    rv4 = 2
                                 End If
                             End If
                         Else
                         rv4 = 4
                         End If
                         Call LabelMenu(3, rv4, rv3)
                 End If
                
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
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
                        ElseIf rv2 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv2 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                         ElseIf rv3 = WRITE_FAIL Or rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv3 = READ_FAIL Or rv4 = READ_FAIL Then
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
Public Sub MultiSlotTestAU6376ALF22()
' add XD MS data pin bonding error sorting
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
 If ChipName = "AU6370DLF20" Or ChipName = "AU6378ALF20" Then
     TmpChip = ChipName
     ChipName = "AU6376"
 End If
            
                'GPIO control setting
If ChipName = "AU6370BL" Or InStr(ChipName, "AU6375HL") <> 0 Or ChipName = "AU6375CL" Or ChipName = "AU6377ALF21" Or ChipName = "AU6377ALS10" Then
     TmpChip = ChipName
     ChipName = "AU6376"
End If
                
If ChipName = "AU6376ELF22" Or ChipName = "AU6376ILF20" Then
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
                  
 
                  
 If TmpChip = "AU6378ALF20" Then
                    
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
                If TmpChip = "AU6377ALF25" Then
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
                 If rv3 = 1 Then
                    rv3 = SetOverCurrent(rv3)
                    If rv3 <> 1 Then
                    rv3 = 2
                    End If
                 End If
                 
                 
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
                
                
                If TmpChip = "AU6378ALF20" Then
                
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
                   '
                If TestResult = "PASS" Then
                  TestResult = ""
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
                   
                    If TmpChip = "AU6378ALF20" Then
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
                 
                 If rv0 = 1 And Left(TmpChip, 10) = "AU6375HLF2" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                    
                    ' AU6375 ram unstable
                    
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 5
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6375(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                    
                    
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
               
                If rv2 = 1 And TmpChip = "AU6378ALF20" Then         '--- for SMC
                
                CardResult = DO_WritePort(card, Channel_P1A, &H18)  ' 0110 0100
                Call MsecDelay(0.5)
                ClosePipe
                rv2 = CBWTest_New(2, rv2, VidName)
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
                 End If
                 
              If rv2 = 1 And Left(TmpChip, 10) = "AU6377ALF2" And TmpChip <> "AU6377ALF21" Then           '--- for SMC
                
                CardResult = DO_WritePort(card, Channel_P1A, &H8 + RomSelector) ' 0110 0100
                Call MsecDelay(0.5)
                ClosePipe
                rv2 = CBWTest_New(2, rv2, VidName)
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
              End If
                
                
               If rv2 = 1 And (TmpChip = "AU6376JLF20") Then      '--- for SMC
                
                  CardResult = DO_WritePort(card, Channel_P1A, &H18)   ' 0110 0100
                  Call MsecDelay(0.5)
                  CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                  Call MsecDelay(0.5)
                 ClosePipe
                 rv2 = CBWTest_New(2, rv2, VidName)
                 Call LabelMenu(2, rv2, rv1)
                 ClosePipe
               End If
               
               
                CardResult = DO_WritePort(card, Channel_P1A, &H18)   ' 0110 0100
                  Call MsecDelay(0.5)
                  CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                  Call MsecDelay(0.5)
                 ClosePipe
                 rv2 = CBWTest_New(2, rv2, VidName)
                 Call LabelMenu(2, rv2, rv1)
                 ClosePipe
               
              
          
               '=============== SMC test END ==================================================
               
               rv3 = CBWTest_New(3, rv2, VidName)  ' MS test
               ClosePipe
               Call LabelMenu(3, rv3, rv2)
             '========================================================
             
                  CardResult = DO_WritePort(card, Channel_P1A, &H18)   ' 0110 0100
                  Call MsecDelay(0.5)
                  CardResult = DO_WritePort(card, Channel_P1A, &H10)   ' 0110 0100
                  Call MsecDelay(0.5)
           
                  rv2 = CBWTest_New(2, rv2, VidName)
                  If rv2 = 1 Then
                   rv2 = Read_OverCurrent(0, 0, 64)
                   If rv2 <> 1 Then
                   rv2 = 2
                   End If
                 End If
                   
                  Call LabelMenu(2, rv2, rv1)
                  ClosePipe
           
             
               If TmpChip = "AU6375HLF21" Then
               
                 If rv0 = 1 Then
                   
                    ClosePipe
                     rv0 = CBWTest_New_AU6375IncPattern(0, 1, VidName)
                     Call LabelMenu(0, rv0, 1)
                     ClosePipe
                 End If
                
                End If
                
                 
                 
                If Left(TmpChip, 10) = "AU6377ALF2" Then
                
                ' test chip
                         ClosePipe
                         rv4 = CBWTest_New(4, rv3, VidName)   'MMC test
                          Call LabelMenu(10, rv4, rv3)
                          ClosePipe
          
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If LightOff <> 127 Then
                            UsbSpeedTestResult = GPO_FAIL
                            rv4 = 2
                         End If
          
                End If
                 
                If TmpChip = "AU6378ALF20" Then
                
             
                         ClosePipe
                         rv4 = CBWTest_New(4, rv3, VidName)
                          Call LabelMenu(10, rv4, rv3)
                          ClosePipe
          
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If rv4 = 1 And LightOff <> 252 Then
                            UsbSpeedTestResult = GPO_FAIL
                            rv4 = 2
                         End If
          
                End If
                 
                 
                    
                  If ChipName = "AU6376" And TmpChip = "AU6370DLF20" Then
                  Call MsecDelay(0.1)
                  rv4 = 1
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 If LightOff <> 254 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                 End If
                     Call LabelMenu(3, rv4, rv3)
                     
                        
                 End If
                 
                 
                 If ChipName = "AU6368A1" Then
                 Call MsecDelay(0.1)
                  rv4 = 1
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 If LightOff <> 192 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                 End If
                     Call LabelMenu(3, rv4, rv3)
                     
                       CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 
                 End If
                 
                  If ChipName = "AU6376" And (Left(TmpChip, 10) <> "AU6377ALF2" And TmpChip <> "AU6378ALF20") Then
                  Call MsecDelay(0.1)
                  rv4 = 1
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                               If TmpChip = "AU6370DLF20" Or TmpChip = "AU6376JLF20" Then
                               
                                  If LightOff <> 254 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                 End If
                               Else
                                 If LightOff <> 252 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                 End If
                              End If
                     Call LabelMenu(3, rv4, rv3)
                     
                        
                 End If
                 
                If ChipName = "AU6368A" Then
                    If rv3 = 1 Then
                           CardResult = DO_WritePort(card, Channel_P1A, &H74)  ' 0111 0100
                           Call MsecDelay(0.1)
                           CardResult = DO_WritePort(card, Channel_P1A, &H54)  ' 0101 0100
                           Call MsecDelay(0.1)
                           rv4 = CBWTest_New(3, rv3, VidName)
                             ClosePipe
                           If rv4 = 1 Then
                                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 If LightOff <> 132 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                    rv4 = 2
                                 End If
                             End If
                         Else
                         rv4 = 4
                         End If
                         Call LabelMenu(3, rv4, rv3)
                 End If
                
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
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
                        ElseIf rv2 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv2 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                         ElseIf rv3 = WRITE_FAIL Or rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv3 = READ_FAIL Or rv4 = READ_FAIL Then
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
   
  Public Sub MultiSlotTestAU6376ALO11()
' add XD MS data pin bonding error sorting
Dim TmpChip As String
Dim RomSelector As Byte
               
  Call PowerSet2(1, "3.08", "0.5", 1, "3.08", "0.5", 1)
  If ChipName = "AU6370GLF20" Then
      ChipName = "AU6370DLF20"
  End If
                
                ' open power
 If ChipName = "AU6377ALF24" Or ChipName = "AU6377ALF25" Then
     TmpChip = ChipName
     ChipName = "AU6376"
 End If
                
                
            '    PowerSet (1) ' for 3.3V , 2.5 V
 If ChipName = "AU6370DLF20" Or ChipName = "AU6378ALF20" Then
     TmpChip = ChipName
     ChipName = "AU6376"
 End If
            
                'GPIO control setting
If ChipName = "AU6370BL" Or InStr(ChipName, "AU6375HL") <> 0 Or ChipName = "AU6375CL" Or ChipName = "AU6377ALF21" Or ChipName = "AU6377ALS10" Then
     TmpChip = ChipName
     ChipName = "AU6376"
End If
                
If ChipName = "AU6376ELF22" Or ChipName = "AU6376ILF20" Then
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
                  
 
                  
 If TmpChip = "AU6378ALF20" Then
                    
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
                If TmpChip = "AU6377ALF25" Then
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
                
                
                If TmpChip = "AU6378ALF20" Then
                
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
                   '
                If TestResult = "PASS" Then
                  TestResult = ""
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
                   
                    If TmpChip = "AU6378ALF20" Then
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
                 
                 If rv0 = 1 And Left(TmpChip, 10) = "AU6375HLF2" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                    
                    ' AU6375 ram unstable
                    
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 5
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6375(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                    
                    
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
               
                If rv2 = 1 And TmpChip = "AU6378ALF20" Then         '--- for SMC
                
                CardResult = DO_WritePort(card, Channel_P1A, &H18)  ' 0110 0100
                Call MsecDelay(0.5)
                ClosePipe
                rv2 = CBWTest_New(2, rv2, VidName)
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
                 End If
                 
              If rv2 = 1 And Left(TmpChip, 10) = "AU6377ALF2" And TmpChip <> "AU6377ALF21" Then           '--- for SMC
                
                CardResult = DO_WritePort(card, Channel_P1A, &H8 + RomSelector) ' 0110 0100
                Call MsecDelay(0.5)
                ClosePipe
                rv2 = CBWTest_New(2, rv2, VidName)
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
              End If
                
                
               If rv2 = 1 And (TmpChip = "AU6376JLF20") Then      '--- for SMC
                
                  CardResult = DO_WritePort(card, Channel_P1A, &H18)   ' 0110 0100
                  Call MsecDelay(0.5)
                  CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                  Call MsecDelay(0.5)
                 ClosePipe
                 rv2 = CBWTest_New(2, rv2, VidName)
                 Call LabelMenu(2, rv2, rv1)
                 ClosePipe
               End If
               
               
                CardResult = DO_WritePort(card, Channel_P1A, &H18)   ' 0110 0100
                  Call MsecDelay(0.5)
                  CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                  Call MsecDelay(0.5)
                 ClosePipe
                 rv2 = CBWTest_New(2, rv2, VidName)
                 Call LabelMenu(2, rv2, rv1)
                 ClosePipe
               
              
          
               '=============== SMC test END ==================================================
               
               rv3 = CBWTest_New(3, rv2, VidName)  ' MS test
               ClosePipe
               Call LabelMenu(3, rv3, rv2)
             '========================================================
             
                  CardResult = DO_WritePort(card, Channel_P1A, &H18)   ' 0110 0100
                  Call MsecDelay(0.5)
                  CardResult = DO_WritePort(card, Channel_P1A, &H10)   ' 0110 0100
                  Call MsecDelay(0.5)
           
                  rv2 = CBWTest_New(2, rv2, VidName)
                
                  Call LabelMenu(2, rv2, rv1)
                  ClosePipe
           
             
               If TmpChip = "AU6375HLF21" Then
               
                 If rv0 = 1 Then
                   
                    ClosePipe
                     rv0 = CBWTest_New_AU6375IncPattern(0, 1, VidName)
                     Call LabelMenu(0, rv0, 1)
                     ClosePipe
                 End If
                
                End If
                
                 
                 
                If Left(TmpChip, 10) = "AU6377ALF2" Then
                
                ' test chip
                         ClosePipe
                         rv4 = CBWTest_New(4, rv3, VidName)   'MMC test
                          Call LabelMenu(10, rv4, rv3)
                          ClosePipe
          
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If LightOff <> 127 Then
                            UsbSpeedTestResult = GPO_FAIL
                            rv4 = 2
                         End If
          
                End If
                 
                If TmpChip = "AU6378ALF20" Then
                
             
                         ClosePipe
                         rv4 = CBWTest_New(4, rv3, VidName)
                          Call LabelMenu(10, rv4, rv3)
                          ClosePipe
          
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If rv4 = 1 And LightOff <> 252 Then
                            UsbSpeedTestResult = GPO_FAIL
                            rv4 = 2
                         End If
          
                End If
                 
                 
                    
                  If ChipName = "AU6376" And TmpChip = "AU6370DLF20" Then
                  Call MsecDelay(0.1)
                  rv4 = 1
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 If LightOff <> 254 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                 End If
                     Call LabelMenu(3, rv4, rv3)
                     
                        
                 End If
                 
                 
                 If ChipName = "AU6368A1" Then
                 Call MsecDelay(0.1)
                  rv4 = 1
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 If LightOff <> 192 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                 End If
                     Call LabelMenu(3, rv4, rv3)
                     
                       CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 
                 End If
                 
                  If ChipName = "AU6376" And (Left(TmpChip, 10) <> "AU6377ALF2" And TmpChip <> "AU6378ALF20") Then
                  Call MsecDelay(0.1)
                  rv4 = 1
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                               If TmpChip = "AU6370DLF20" Or TmpChip = "AU6376JLF20" Then
                               
                                  If LightOff <> 254 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                 End If
                               Else
                                 If LightOff <> 252 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                 End If
                              End If
                     Call LabelMenu(3, rv4, rv3)
                     
                        
                 End If
                 
                If ChipName = "AU6368A" Then
                    If rv3 = 1 Then
                           CardResult = DO_WritePort(card, Channel_P1A, &H74)  ' 0111 0100
                           Call MsecDelay(0.1)
                           CardResult = DO_WritePort(card, Channel_P1A, &H54)  ' 0101 0100
                           Call MsecDelay(0.1)
                           rv4 = CBWTest_New(3, rv3, VidName)
                             ClosePipe
                           If rv4 = 1 Then
                                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 If LightOff <> 132 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                    rv4 = 2
                                 End If
                             End If
                         Else
                         rv4 = 4
                         End If
                         Call LabelMenu(3, rv4, rv3)
                 End If
                
                  OpenPipe
                  If rv4 = 1 Then
                   
                    If rv4 = 1 Then
                     rv5 = SetOverCurrent(rv4)
                        If rv5 <> 1 Then
                          rv5 = 2
                        End If
                     End If
                     
                   If rv5 = 1 Then
                        rv5 = Read_OverCurrent(0, 0, 64)
                        If rv5 <> 1 Then
                        rv5 = 2
                        End If
                   End If
                   
                 End If
                    ClosePipe
                    Call LabelMenu(51, rv5, rv4)
                    If rv5 <> 1 Then
                    Tester.Label9.Caption = "Over Current fail"
                    Tester.Label2.Caption = "Over Current fail---"
                    End If
                   
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv5, " \\OverCurret :0 Fail, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
               
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
                        ElseIf rv2 = WRITE_FAIL Or rv3 = WRITE_FAIL Or rv4 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv2 = READ_FAIL Or rv3 = READ_FAIL Or rv4 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                         ElseIf rv5 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv5 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                            
                            
                        ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
                
               
                  
                End If
                CardResult = DO_WritePort(card, Channel_P1A, &H1)
                  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
   End Sub
   
   Public Sub MultiSlotTestAU6376ALO10()
' add XD MS data pin bonding error sorting
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
 If ChipName = "AU6370DLF20" Or ChipName = "AU6378ALF20" Then
     TmpChip = ChipName
     ChipName = "AU6376"
 End If
            
                'GPIO control setting
If ChipName = "AU6370BL" Or InStr(ChipName, "AU6375HL") <> 0 Or ChipName = "AU6375CL" Or ChipName = "AU6377ALF21" Or ChipName = "AU6377ALS10" Then
     TmpChip = ChipName
     ChipName = "AU6376"
End If
                
If ChipName = "AU6376ELF22" Or ChipName = "AU6376ILF20" Then
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
                  
 
                  
 If TmpChip = "AU6378ALF20" Then
                    
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
                If TmpChip = "AU6377ALF25" Then
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
                
                
                If TmpChip = "AU6378ALF20" Then
                
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
                   '
                If TestResult = "PASS" Then
                  TestResult = ""
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
                   
                    If TmpChip = "AU6378ALF20" Then
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
                 
                 If rv0 = 1 And Left(TmpChip, 10) = "AU6375HLF2" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                    
                    ' AU6375 ram unstable
                    
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 5
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6375(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                    
                    
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
               
                If rv2 = 1 And TmpChip = "AU6378ALF20" Then         '--- for SMC
                
                CardResult = DO_WritePort(card, Channel_P1A, &H18)  ' 0110 0100
                Call MsecDelay(0.5)
                ClosePipe
                rv2 = CBWTest_New(2, rv2, VidName)
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
                 End If
                 
              If rv2 = 1 And Left(TmpChip, 10) = "AU6377ALF2" And TmpChip <> "AU6377ALF21" Then           '--- for SMC
                
                CardResult = DO_WritePort(card, Channel_P1A, &H8 + RomSelector) ' 0110 0100
                Call MsecDelay(0.5)
                ClosePipe
                rv2 = CBWTest_New(2, rv2, VidName)
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
              End If
                
                
               If rv2 = 1 And (TmpChip = "AU6376JLF20") Then      '--- for SMC
                
                  CardResult = DO_WritePort(card, Channel_P1A, &H18)   ' 0110 0100
                  Call MsecDelay(0.5)
                  CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                  Call MsecDelay(0.5)
                 ClosePipe
                 rv2 = CBWTest_New(2, rv2, VidName)
                 Call LabelMenu(2, rv2, rv1)
                 ClosePipe
               End If
               
               
                CardResult = DO_WritePort(card, Channel_P1A, &H18)   ' 0110 0100
                  Call MsecDelay(0.5)
                  CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                  Call MsecDelay(0.5)
                 ClosePipe
                 rv2 = CBWTest_New(2, rv2, VidName)
                 Call LabelMenu(2, rv2, rv1)
                 ClosePipe
               
              
          
               '=============== SMC test END ==================================================
               
               rv3 = CBWTest_New(3, rv2, VidName)  ' MS test
               ClosePipe
               Call LabelMenu(3, rv3, rv2)
             '========================================================
             
                  CardResult = DO_WritePort(card, Channel_P1A, &H18)   ' 0110 0100
                  Call MsecDelay(0.5)
                  CardResult = DO_WritePort(card, Channel_P1A, &H10)   ' 0110 0100
                  Call MsecDelay(0.5)
           
                  rv2 = CBWTest_New(2, rv2, VidName)
                
                  Call LabelMenu(2, rv2, rv1)
                  ClosePipe
           
             
               If TmpChip = "AU6375HLF21" Then
               
                 If rv0 = 1 Then
                   
                    ClosePipe
                     rv0 = CBWTest_New_AU6375IncPattern(0, 1, VidName)
                     Call LabelMenu(0, rv0, 1)
                     ClosePipe
                 End If
                
                End If
                
                 
                 
                If Left(TmpChip, 10) = "AU6377ALF2" Then
                
                ' test chip
                         ClosePipe
                         rv4 = CBWTest_New(4, rv3, VidName)   'MMC test
                          Call LabelMenu(10, rv4, rv3)
                          ClosePipe
          
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If LightOff <> 127 Then
                            UsbSpeedTestResult = GPO_FAIL
                            rv4 = 2
                         End If
          
                End If
                 
                If TmpChip = "AU6378ALF20" Then
                
             
                         ClosePipe
                         rv4 = CBWTest_New(4, rv3, VidName)
                          Call LabelMenu(10, rv4, rv3)
                          ClosePipe
          
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If rv4 = 1 And LightOff <> 252 Then
                            UsbSpeedTestResult = GPO_FAIL
                            rv4 = 2
                         End If
          
                End If
                 
                 
                    
                  If ChipName = "AU6376" And TmpChip = "AU6370DLF20" Then
                  Call MsecDelay(0.1)
                  rv4 = 1
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 If LightOff <> 254 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                 End If
                     Call LabelMenu(3, rv4, rv3)
                     
                        
                 End If
                 
                 
                 If ChipName = "AU6368A1" Then
                 Call MsecDelay(0.1)
                  rv4 = 1
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 If LightOff <> 192 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                 End If
                     Call LabelMenu(3, rv4, rv3)
                     
                       CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 
                 End If
                 
                  If ChipName = "AU6376" And (Left(TmpChip, 10) <> "AU6377ALF2" And TmpChip <> "AU6378ALF20") Then
                  Call MsecDelay(0.1)
                  rv4 = 1
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                               If TmpChip = "AU6370DLF20" Or TmpChip = "AU6376JLF20" Then
                               
                                  If LightOff <> 254 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                 End If
                               Else
                                 If LightOff <> 252 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                 End If
                              End If
                     Call LabelMenu(3, rv4, rv3)
                     
                        
                 End If
                 
                If ChipName = "AU6368A" Then
                    If rv3 = 1 Then
                           CardResult = DO_WritePort(card, Channel_P1A, &H74)  ' 0111 0100
                           Call MsecDelay(0.1)
                           CardResult = DO_WritePort(card, Channel_P1A, &H54)  ' 0101 0100
                           Call MsecDelay(0.1)
                           rv4 = CBWTest_New(3, rv3, VidName)
                             ClosePipe
                           If rv4 = 1 Then
                                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 If LightOff <> 132 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                    rv4 = 2
                                 End If
                             End If
                         Else
                         rv4 = 4
                         End If
                         Call LabelMenu(3, rv4, rv3)
                 End If
                
                  OpenPipe
                  If rv4 = 1 Then
                   
                    If rv4 = 1 Then
                     rv5 = SetOverCurrent(rv4)
                        If rv5 <> 1 Then
                          rv5 = 2
                        End If
                     End If
                     
                   If rv5 = 1 Then
                        rv5 = Read_OverCurrent(0, 0, 64)
                        If rv5 <> 1 Then
                        rv5 = 2
                        End If
                   End If
                   
                 End If
                    ClosePipe
                    Call LabelMenu(51, rv5, rv4)
                    If rv5 <> 1 Then
                    Tester.Label9.Caption = "Over Current fail"
                    Tester.Label2.Caption = "Over Current fail---"
                    End If
                   
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv4, " \\MSPro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv5, " \\OverCurret :0 Fail, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
               
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
                        ElseIf rv2 = WRITE_FAIL Or rv3 = WRITE_FAIL Or rv4 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv2 = READ_FAIL Or rv3 = READ_FAIL Or rv4 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                         ElseIf rv5 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv5 = READ_FAIL Then
                            MSReadFail = MSReadFail + 1
                            TestResult = "MS_RF"
                            
                            
                        ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
                
               
                  
                End If
                CardResult = DO_WritePort(card, Channel_P1A, &H1)
                  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
   End Sub
   Public Sub MultiSlotTestAU6376ALF21()
' add XD MS data pin bonding error sorting
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
 If ChipName = "AU6370DLF20" Or ChipName = "AU6378ALF20" Then
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
                  
 
                  
 If TmpChip = "AU6378ALF20" Then
                    
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
                If TmpChip = "AU6377ALF25" Then
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
                
                
                If TmpChip = "AU6378ALF20" Then
                
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
                   '
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
                   
                    If TmpChip = "AU6378ALF20" Then
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
                 
                 If rv0 = 1 And Left(TmpChip, 10) = "AU6375HLF2" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                    ClosePipe
                    
                    ' AU6375 ram unstable
                    
                    TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 5
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6375(0, 1)  ' write
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                    
                    
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
               
                If rv2 = 1 And TmpChip = "AU6378ALF20" Then         '--- for SMC
                
                CardResult = DO_WritePort(card, Channel_P1A, &H18)  ' 0110 0100
                Call MsecDelay(0.5)
                ClosePipe
                rv2 = CBWTest_New(2, rv2, VidName)
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
                 End If
                 
              If rv2 = 1 And Left(TmpChip, 10) = "AU6377ALF2" And TmpChip <> "AU6377ALF21" Then           '--- for SMC
                
                CardResult = DO_WritePort(card, Channel_P1A, &H8 + RomSelector) ' 0110 0100
                Call MsecDelay(0.5)
                ClosePipe
                rv2 = CBWTest_New(2, rv2, VidName)
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
              End If
                
                
               If rv2 = 1 And (TmpChip = "AU6376JLF20") Then      '--- for SMC
                
                  CardResult = DO_WritePort(card, Channel_P1A, &H18)   ' 0110 0100
                  Call MsecDelay(0.5)
                  CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                  Call MsecDelay(0.5)
                 ClosePipe
                 rv2 = CBWTest_New(2, rv2, VidName)
                 Call LabelMenu(2, rv2, rv1)
                 ClosePipe
               End If
               
               
                CardResult = DO_WritePort(card, Channel_P1A, &H18)   ' 0110 0100
                  Call MsecDelay(0.5)
                  CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 0110 0100
                  Call MsecDelay(0.5)
                 ClosePipe
                 rv2 = CBWTest_New(2, rv2, VidName)
                 Call LabelMenu(2, rv2, rv1)
                 ClosePipe
               
              
          
               '=============== SMC test END ==================================================
               
               rv3 = CBWTest_New(3, rv2, VidName)  ' MS test
               ClosePipe
               Call LabelMenu(3, rv3, rv2)
             '========================================================
             
                  CardResult = DO_WritePort(card, Channel_P1A, &H18)   ' 0110 0100
                  Call MsecDelay(0.5)
                  CardResult = DO_WritePort(card, Channel_P1A, &H10)   ' 0110 0100
                  Call MsecDelay(0.5)
           
                  rv2 = CBWTest_New(2, rv2, VidName)
                  Call LabelMenu(2, rv2, rv1)
                  ClosePipe
           
             
               If TmpChip = "AU6375HLF21" Then
               
                 If rv0 = 1 Then
                   
                    ClosePipe
                     rv0 = CBWTest_New_AU6375IncPattern(0, 1, VidName)
                     Call LabelMenu(0, rv0, 1)
                     ClosePipe
                 End If
                
                End If
                
                 
                 
                If Left(TmpChip, 10) = "AU6377ALF2" Then
                
                ' test chip
                         ClosePipe
                         rv4 = CBWTest_New(4, rv3, VidName)   'MMC test
                          Call LabelMenu(10, rv4, rv3)
                          ClosePipe
          
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If LightOff <> 127 Then
                            UsbSpeedTestResult = GPO_FAIL
                            rv4 = 2
                         End If
          
                End If
                 
                If TmpChip = "AU6378ALF20" Then
                
             
                         ClosePipe
                         rv4 = CBWTest_New(4, rv3, VidName)
                          Call LabelMenu(10, rv4, rv3)
                          ClosePipe
          
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If rv4 = 1 And LightOff <> 252 Then
                            UsbSpeedTestResult = GPO_FAIL
                            rv4 = 2
                         End If
          
                End If
                 
                 
                    
                  If ChipName = "AU6376" And TmpChip = "AU6370DLF20" Then
                  Call MsecDelay(0.1)
                  rv4 = 1
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 If LightOff <> 254 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                 End If
                     Call LabelMenu(3, rv4, rv3)
                     
                        
                 End If
                 
                 
                 If ChipName = "AU6368A1" Then
                 Call MsecDelay(0.1)
                  rv4 = 1
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 If LightOff <> 192 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                 End If
                     Call LabelMenu(3, rv4, rv3)
                     
                       CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 
                 End If
                 
                  If ChipName = "AU6376" And (Left(TmpChip, 10) <> "AU6377ALF2" And TmpChip <> "AU6378ALF20") Then
                  Call MsecDelay(0.1)
                  rv4 = 1
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                               If TmpChip = "AU6370DLF20" Or TmpChip = "AU6376JLF20" Then
                               
                                  If LightOff <> 254 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                 End If
                               Else
                                 If LightOff <> 252 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                 End If
                              End If
                     Call LabelMenu(3, rv4, rv3)
                     
                        
                 End If
                 
                If ChipName = "AU6368A" Then
                    If rv3 = 1 Then
                           CardResult = DO_WritePort(card, Channel_P1A, &H74)  ' 0111 0100
                           Call MsecDelay(0.1)
                           CardResult = DO_WritePort(card, Channel_P1A, &H54)  ' 0101 0100
                           Call MsecDelay(0.1)
                           rv4 = CBWTest_New(3, rv3, VidName)
                             ClosePipe
                           If rv4 = 1 Then
                                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 If LightOff <> 132 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                    rv4 = 2
                                 End If
                             End If
                         Else
                         rv4 = 4
                         End If
                         Call LabelMenu(3, rv4, rv3)
                 End If
                
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
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
                        ElseIf rv2 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv2 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                         ElseIf rv3 = WRITE_FAIL Or rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv3 = READ_FAIL Or rv4 = READ_FAIL Then
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
Function CBWTest_New_21_Sector_AU6377(Lun As Byte, PreSlotStatus As Byte) As Byte
Dim i As Long
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long
Dim j As Integer

 CBWDataTransferLength = 10752  ' 64 sector
'CBWDataTransferLength = 10240
 'CBWDataTransferLength = 5120
    If PreSlotStatus <> 1 Then
        CBWTest_New_21_Sector_AU6377 = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New_21_Sector_AU6377 = 2
   
    '========================================
    
    
     If OpenPipe = 0 Then
       CBWTest_New_21_Sector_AU6377 = 2   ' Write fail
       Exit Function
     End If
  
    '====================================
     TmpInteger = TestUnitSpeed(Lun)
    
    If TmpInteger = 0 Then
        
       CBWTest_New_21_Sector_AU6377 = 2   ' usb 2.0 high speed fail
       UsbSpeedTestResult = 2
       Exit Function
    End If
    TmpInteger = 0
    
    TmpInteger = TestUnitReady(Lun)
     If TmpInteger = 0 Then
         TmpInteger = RequestSense(Lun)
        
         If TmpInteger = 0 Then
        
            CBWTest_New_21_Sector_AU6377 = 2  'Write fail
            Exit Function
         End If
        
     End If
  
  
   
       
       ' For i = 0 To CBWDataTransferLength - 1
        
       '      ReadData(i) = 0
    
       ' Next

        
     
        TmpInteger = Write_Data_AU6377(LBA, Lun, CBWDataTransferLength)
         
        If TmpInteger = 0 Then
            CBWTest_New_21_Sector_AU6377 = 2  'write fail
         
            Exit Function
        End If
   
        TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
         
        If TmpInteger = 0 Then
            CBWTest_New_21_Sector_AU6377 = 3    'Read fail
             
            Exit Function
        End If
     
        For i = 0 To CBWDataTransferLength - 1
        
            If ReadData(i) <> Pattern_AU6377(i) Then
              CBWTest_New_21_Sector_AU6377 = 3    'Read fail
           
              Exit Function
            End If
        
        Next
   
        
       CBWTest_New_21_Sector_AU6377 = 1
           
         
    End Function
Public Function CBWTest_New_AU6375IncPattern2(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

   CBWDataTransferLength = 4096
 
   For i = 0 To CBWDataTransferLength - 1
    
         IncPattern(i) = i Mod 255
        ' Debug.Print Pattern(i)

    Next

    If PreSlotStatus <> 1 Then
        CBWTest_New_AU6375IncPattern2 = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New_AU6375IncPattern2 = 0
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
      CBWTest_New_AU6375IncPattern2 = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_New_AU6375IncPattern2 = 2   ' Write fail
      Exit Function
    End If
 
    '======================================
    
     ' for unitSpeed
    
   
    
    
    
    TmpInteger = TestUnitReady(Lun)
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
           CBWTest_New_AU6375IncPattern2 = 2  'Write fail
           Exit Function
        End If
        
    End If
    '======================================
    If ChipName = "AU6371" Then
        TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
    End If
    
   ' TmpInteger = Read_Data1(Lba, Lun, CBWDataTransferLength)
    TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
      
    If TmpInteger = 0 Then
         CBWTest_New_AU6375IncPattern2 = 2  'write fail
          Exit Function
     End If
    
      
    TmpInteger = Write_DataIncPattern(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_AU6375IncPattern2 = 2  'write fail
        Exit Function
    End If
    
    TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_AU6375IncPattern2 = 3    'Read fail
        Exit Function
    End If
     
    For i = 0 To CBWDataTransferLength - 1
    
        If ReadData(i) <> IncPattern(i) Then
          CBWTest_New_AU6375IncPattern2 = 3    'Read fail
          Exit Function
        End If
    
    Next
    
    CBWTest_New_AU6375IncPattern2 = 1
        
    
    End Function

Public Sub AU6433F23MPTest()

If ChipName = "AU6433EFF23" Then
    Call AU6433EFF23TestSub

ElseIf ChipName = "AU6433BLF24" Then
    Call AU6433BLF24TestSub

ElseIf ChipName = "AU6433BLF25" Then
    Call AU6433BLF25TestSub

ElseIf ChipName = "AU6433BLF26" Then
    Call AU6433BLF26TestSub

ElseIf ChipName = "AU6433BLF27" Then
    Call AU6433BLF27TestSub

ElseIf ChipName = "AU6433BLF28" Then
    Call AU6433BLF28TestSub

ElseIf ChipName = "AU6433BLF29" Then
    Call AU6433BLF29TestSub

ElseIf ChipName = "AU6433BLF2E" Then
    Call AU6433BLF2ETestSub
    
ElseIf ChipName = "AU6433BLF2F" Then
    Call AU6433BLF2FTestSub
    
ElseIf ChipName = "AU6433BLFEE" Then
    Call AU6433BLFEETestSub

ElseIf ChipName = "AU6433BLF3B" Then
    Call AU6433BLF3BTestSub

ElseIf ChipName = "AU6433BLF3C" Then
    Call AU6433BLF3CTestSub
    
ElseIf ChipName = "AU6433BLF3D" Then
    Call AU6433BLF3DTestSub

ElseIf ChipName = "AU6433BLF3E" Then
    Call AU6433BLF3ETestSub
    
ElseIf ChipName = "AU6433BLF3F" Then
    Call AU6433BLF3FTestSub

ElseIf ChipName = "AU6433BLF0E" Then
    Call AU6433BLF0ETestSub
    
ElseIf ChipName = "AU6433BLF0F" Then
    Call AU6433BLF0FTestSub

ElseIf ChipName = "AU6433BLD3B" Then
    Call AU6433BLD3BTestSub
    
ElseIf ChipName = "AU6433BLF2A" Then
    Call AU6433BLF2ATestSub

ElseIf ChipName = "AU6433BLF2B" Then
    Call AU6433BLF2ATestSub

ElseIf ChipName = "AU6433BLF2C" Then
    Call AU6433BLF2CTestSub
    
ElseIf ChipName = "AU6433BLF2D" Then
    Call AU6433BLF2DTestSub
    
ElseIf ChipName = "AU6433BLF30" Then
    Call AU6433BLF30TestSub
    
ElseIf ChipName = "AU6433CSF28" Then
    Call AU6433CSF28TestSub
    
ElseIf ChipName = "AU6433CSF29" Then
    Call AU6433CSF29TestSub

ElseIf ChipName = "AU6433CSF2A" Then
    Call AU6433CSF2ATestSub

ElseIf ChipName = "AU6433CSF09" Then
    Call AU6433CSF09TestSub
    
ElseIf ChipName = "AU6433CSF0A" Then
    Call AU6433CSF0ATestSub

ElseIf ChipName = "AU6433DLF20" Then
    Call AU6433DLF20TestSub

ElseIf ChipName = "AU6433DLF21" Then
    Call AU6433DLF21TestSub

ElseIf ChipName = "AU6433DLF22" Then
    Call AU6433DLF22TestSub

ElseIf ChipName = "AU6433DLF23" Then
    Call AU6433DLF23TestSub
    
ElseIf ChipName = "AU6433DLF03" Then
    Call AU6433DLF03TestSub

ElseIf ChipName = "AU6433DLF30" Then
    Call AU6433DLF30TestSub

ElseIf ChipName = "AU6433DLF31" Then
    Call AU6433DLF31TestSub

ElseIf ChipName = "AU6433DLF32" Then
    Call AU6433DLF32TestSub
    
ElseIf ChipName = "AU6433DLF33" Then
    Call AU6433DLF33TestSub
    
ElseIf ChipName = "AU6433DLF00" Then
    Call AU6433DLF00TestSub

ElseIf ChipName = "AU6433DLF3C" Then
    Call AU6433DLF3CTestSub

ElseIf ChipName = "AU6433FSF28" Then
    Call AU6433FSF28TestSub
    
ElseIf ChipName = "AU6433FSF29" Then
    Call AU6433FSF29TestSub
    
ElseIf ChipName = "AU6433JSF28" Then
    Call AU6433JSF28TestSub

ElseIf ChipName = "AU6433JSF29" Then
    Call AU6433JSF29TestSub
    
ElseIf ChipName = "AU6433JSF39" Then
    Call AU6433JSF39TestSub

ElseIf ChipName = "AU6433EFF35" Then
    Call AU6433EFF35TestSub

ElseIf ChipName = "AU6433EFF36" Then
    Call AU6433EFF36TestSub

ElseIf ChipName = "AU6433EFF3F" Then
    Call AU6433EFF36TestSub

ElseIf ChipName = "AU6433EFF25" Then
    Call AU6433EFF25TestSub

ElseIf ChipName = "AU6433EFF26" Then
    Call AU6433EFF26TestSub

ElseIf ChipName = "AU6433BLS10" Then
    Call AU6433BLS10SortingSub

ElseIf ChipName = "AU6433BLS11" Then
    Call AU6433BLS11SortingSub

ElseIf ChipName = "AU6433BLS12" Then
    Call AU6433BLS12SortingSub

ElseIf ChipName = "AU6433BLS13" Then
    Call AU6433BLS13SortingSub

ElseIf ChipName = "AU6433LFF33" Then
    Call AU6433LFF33TestSub

ElseIf ChipName = "AU6433LFF23" Then
    Call AU6433LFF23TestSub

ElseIf ChipName = "AU6433KFF23" Then
    Call AU6433KFF23TestSub

ElseIf ChipName = "AU6433GSF23" Then
    Call AU6433GSF23TestSub

ElseIf ChipName = "AU6433DFF23" Then
    Call AU6433DFF23TestSub

ElseIf ChipName = "AU6433HFF23" Then
    Call AU6433HFF23TestSub

ElseIf ChipName = "AU6433HSF23" Then
    Call AU6433HSF23TestSub

ElseIf ChipName = "AU6433ESF23" Then
    Call AU6433ESF23TestSub

ElseIf ChipName = "AU6433FSF23" Then
    Call AU6433FSF23TestSub
ElseIf ChipName = "AU6433IFF23" Then
    Call AU6433IFF23TestSub
ElseIf ChipName = "AU6433VFF23" Then
    Call AU6433VFF23TestSub
End If


End Sub





Public Sub AU6433F22MPTest()

If ChipName = "AU6433EFF22" Then
    Call AU6433EFF22TestSub

End If

If ChipName = "AU6433LFF22" Then
    Call AU6433LFF22TestSub

End If


If ChipName = "AU6433KFF22" Then
    Call AU6433KFF22TestSub

End If

If ChipName = "AU6433GSF22" Then
    Call AU6433GSF22TestSub

End If


If ChipName = "AU6433DFF22" Then
    Call AU6433DFF22TestSub

End If

If ChipName = "AU6433HFF22" Then
    Call AU6433HFF22TestSub

End If

If ChipName = "AU6433HSF22" Then
    Call AU6433HSF22TestSub

End If



If ChipName = "AU6433ESF22" Then
    Call AU6433ESF22TestSub

End If

If ChipName = "AU6433FSF22" Then
    Call AU6433FSF22TestSub
End If

If ChipName = "AU6433IFF22" Then
    Call AU6433IFF22TestSub

End If

End Sub

Public Sub SingleSlotTest()

Dim TmpChip As String

TmpChip = Left(ChipName, 10)
     
Select Case TmpChip

Case "AU6371ELF2", "AU6371GLF2", "AU6371HLF2"
     
        Call AU6371ELTest
     
Case "AU6371DLF2", "AU6371BLF2"
     
       Call AU6371DLTest
       
Case "AU6371AFF2"
    
       Call AU6371AFTest
       
Case "AU6371CFF2"

       Call AU6371CFTest
       
       
Case "AU6366CLF2"
       
       Call AU6371CFTest
       
Case "AU6366ALF2"
       
       Call AU6366ALF20TestSub
       
       
Case "AU6433EFF2"
       
       Call AU6433EFTest
       
Case "AU6433HFF2"
       
       Call AU6433HFTest
       
Case "AU6433DFF2"
       
       Call AU6433DFTest
Case "AU6433BSF2"
       
       Call AU6433BSTest
       
 Case "AU6433KFF2"
       
       Call AU6433KFTest
       
 Case "AU6433GSF2"
       
       Call AU6433GSTest
       
     
End Select
     
End Sub
     

Public Sub SingleSlotTest25()

Dim TmpChip As String

TmpChip = Left(ChipName, 10)
     
Select Case TmpChip

Case "AU6371GLF2", "AU6371HLF2"
     
        Call AU6371ELTest
     
Case "AU6371ELF2"
     
        Call AU6371ELTest25
Case "AU6371DLF2", "AU6371BLF2"
     
       Call AU6371DLTest25
       
Case "AU6371AFF2"
    
       Call AU6371AFTest
       
Case "AU6371CFF2"

       Call AU6371CFTest
       
       
Case "AU6366CLF2"
       
       Call AU6371CFTest
       
     
End Select
     
     
End Sub

Public Sub SingleSlotTest27()

Dim TmpChip As String

TmpChip = Left(ChipName, 10)
     
Select Case TmpChip

Case "AU6371TLF2"
     
        Call AU6371TLTest27

Case "AU6371HLF2"

       Call AU6371HLTest27
Case "AU6371SLF2"
     
        Call AU6371SLTest27

Case "AU6371NLF2"
     
        Call AU6371NLTest26

Case "AU6371GLF2"
     
        Call AU6371GLTest27
  
Case "AU6471FLF2"
        Call AU6471FLTest20
Case "AU6471GLF2"
        Call AU6471GLTest20
Case "AU6371ELF2"
     
        Call AU6371ELTest27
        
Case "AU6371ELS3"
     
        Call AU6371ELNoteBookSorting
        
Case "AU6371DLF2", "AU6371BLF2"
     
       Call AU6371DLTest27
       
Case "AU6371AFF2"
    
       Call AU6371AFTest
       
Case "AU6371CFF2"

       Call AU6371CFTest
       
       
Case "AU6366CLF2"
       
       Call AU6371CFTest
       
       
     
End Select
     
     
End Sub


Public Sub AU6371MP28()

Dim TmpChip As String

TmpChip = Left(ChipName, 10)
     
Select Case TmpChip

Case "AU6371TLF2"
     
        Call AU6371TLTest28

Case "AU6371HLF2"

       Call AU6371HLTest28
Case "AU6371SLF2"
     
        Call AU6371SLTest28

 
Case "AU6371GLF2"
     
        Call AU6371GLTest28
  
 Case "AU6371ELF2"
     
        Call AU6371ELTest28
        
 
        
Case "AU6371DLF2", "AU6371BLF2"
     
       Call AU6371DLTest28
       
 
       
       
     
End Select
     
     
End Sub

Public Sub AU6371MP29()

Dim TmpChip As String

TmpChip = Left(ChipName, 10)
     
Select Case TmpChip

Case "AU6371TLF2"
     
        Call AU6371TLTest29

Case "AU6371HLF2"

       Call AU6371HLTest29
Case "AU6371SLF2"
     
        Call AU6371SLTest29

 
Case "AU6371GLF2"
     
        Call AU6371GLTest29
  
 Case "AU6371ELF2"
     
        Call AU6371ELTest29
        
 
        
Case "AU6371DLF2", "AU6371BLF2"
     
       Call AU6371DLTest29
       
 Case "AU6371DLO1"
       Call AU6371DLTestOverCurrent
       
     
End Select
     
     
End Sub
     
     
 Public Sub SingleSlotTest26()

Dim TmpChip As String



TmpChip = Left(ChipName, 10)
     
Select Case TmpChip

Case "AU6371TLF2"
     
        Call AU6371TLTest26


Case "AU6371SLF2"
     
        Call AU6371SLTest26

Case "AU6371NLF2"
     
        Call AU6371NLTest26

Case "AU6371GLF2", "AU6371HLF2"
     
        Call AU6371ELTest
  
Case "AU6471FLF2"
        Call AU6471FLTest20
Case "AU6471GLF2"
        Call AU6471GLTest20
Case "AU6371ELF2"
     
        Call AU6371ELTest26
Case "AU6371DLF2", "AU6371BLF2"
     
       Call AU6371DLTest26
       
Case "AU6371AFF2"
    
       Call AU6371AFTest
       
Case "AU6371CFF2"

       Call AU6371CFTest
       
       
Case "AU6366CLF2"
       
       Call AU6371CFTest
       
Case "AU6371DLS1"

        Call AU6371DLSorting1
        
Case "AU6371DLS2"

        
        
        If ChipName = "AU6371DLS20" Then
             Call AU6371DLS20SortingSub
        ElseIf ChipName = "AU6371DLS21" Then
             Call AU6371DLS21SortingSub
        ElseIf ChipName = "AU6371DLS22" Then
             Call AU6371DLS22SortingSub
        End If
        
        
Case "AU6371DLS3"
         
        If ChipName = "AU6371DLS30" Then
             Call AU6371DLS30SortingSub
        ElseIf ChipName = "AU6371DLS31" Then
             Call AU6371DLS31SortingSub
          ElseIf ChipName = "AU6371DLS32" Then
             Call AU6371DLS32SortingSub
        End If
        
Case "AU6371DLS4"
         
        If ChipName = "AU6371DLS40" Then
             Call AU6371DLS40SortingSub
        ElseIf ChipName = "AU6371DLS41" Then
             Call AU6371DLS41SortingSub
        
        End If
        
End Select
     
     
End Sub
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
     
              

 




Public Function Write_DataIncPattern(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte

Dim CBW(0 To 30) As Byte
Dim CSW(0 To 12) As Byte
Dim NumberOfBytesWritten As Long
Dim NumberOfBytesRead As Long
Dim CBWDataTransferLen(0 To 3) As Byte
Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim i As Integer
Dim tmpV(0 To 2) As Long
Dim opcode As Byte

opcode = &H2A
'Buffer(0) = &H33 'CByte(Text2.Text)
'Buffer(1) = &H44


    For i = 0 To 30
    
        CBW(i) = 0
    
    Next i
    
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
CBW(12) = &H0                 '80

'////////////// LUN
CBW(13) = Lun                    '00

'///////////// CBD Len
CBW(14) = &HA                '0a

'////////////  UFI command

CBW(15) = opcode
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

 
'1. CBW output
 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)    'out

If result = 0 Then
    Write_DataIncPattern = 0
    Exit Function
End If
 
 
 
'2, Output data
result = WriteFile _
       (WriteHandle, _
       IncPattern(0), _
       CBWDataTransferLength, _
       NumberOfBytesWritten, _
       0)    'out

 
If result = 0 Then
    Write_DataIncPattern = 0
    Exit Function
End If

'3 . CSW
result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
        
If result = 0 Then
    Write_DataIncPattern = 0
    Exit Function
End If
 
 
 
If CSW(12) = 1 Then
Write_DataIncPattern = 0

Else
Write_DataIncPattern = 1
End If
End Function

Public Function CBWTest_New_AU6375IncPattern(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

   CBWDataTransferLength = 4096
 
   For i = 0 To CBWDataTransferLength - 1
    
         IncPattern(i) = i Mod 256
        ' Debug.Print Pattern(i)

    Next

    If PreSlotStatus <> 1 Then
        CBWTest_New_AU6375IncPattern = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New_AU6375IncPattern = 0
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
      CBWTest_New_AU6375IncPattern = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_New_AU6375IncPattern = 2   ' Write fail
      Exit Function
    End If
 
    '======================================
    
     ' for unitSpeed
    
   
    
    
    
    TmpInteger = TestUnitReady(Lun)
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
           CBWTest_New_AU6375IncPattern = 2  'Write fail
           Exit Function
        End If
        
    End If
    '======================================
    If ChipName = "AU6371" Then
        TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
    End If
    
   ' TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
    TmpInteger = Read_Data1(LBA, Lun, 1024)
    TmpInteger = Read_Data1(LBA, Lun, 1024)
    If TmpInteger = 0 Then
         CBWTest_New_AU6375IncPattern = 2  'write fail
          Exit Function
     End If
    
      
    TmpInteger = Write_DataIncPattern(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_AU6375IncPattern = 2  'write fail
        Exit Function
    End If
    
    TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_AU6375IncPattern = 3    'Read fail
        Exit Function
    End If
     
    For i = 0 To CBWDataTransferLength - 1
    
        If ReadData(i) <> IncPattern(i) Then
          CBWTest_New_AU6375IncPattern = 3    'Read fail
          Exit Function
        End If
    
    Next
    
    CBWTest_New_AU6375IncPattern = 1
        
    
    End Function
Public Function CBWTest_New_128_Sector_AU6375(Lun As Byte, PreSlotStatus As Byte) As Byte
Dim i As Long
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long
Dim j As Integer

 CBWDataTransferLength = 65536 ' 64 sector

   
    If PreSlotStatus <> 1 Then
        CBWTest_New_128_Sector_AU6375 = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New_128_Sector_AU6375 = 2
   
    '========================================
    
    
     If OpenPipe = 0 Then
       CBWTest_New_128_Sector_AU6375 = 2   ' Write fail
       Exit Function
     End If
  
    '====================================
     TmpInteger = TestUnitSpeed(Lun)
    
    If TmpInteger = 0 Then
        
       CBWTest_New_128_Sector_AU6375 = 2   ' usb 2.0 high speed fail
       UsbSpeedTestResult = 2
       Exit Function
    End If
    TmpInteger = 0
    
    TmpInteger = TestUnitReady(Lun)
     If TmpInteger = 0 Then
         TmpInteger = RequestSense(Lun)
        
         If TmpInteger = 0 Then
        
            CBWTest_New_128_Sector_AU6375 = 2  'Write fail
            Exit Function
         End If
        
     End If
  
  
   
       
       ' For i = 0 To CBWDataTransferLength - 1
        
       '      ReadData(i) = 0
    
       ' Next

        
     
        TmpInteger = Write_Data_AU6375(LBA, Lun, CBWDataTransferLength)
         
        If TmpInteger = 0 Then
            CBWTest_New_128_Sector_AU6375 = 2  'write fail
         
            Exit Function
        End If
   
        TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
         
        If TmpInteger = 0 Then
            CBWTest_New_128_Sector_AU6375 = 3    'Read fail
             
            Exit Function
        End If
     
        For i = 0 To CBWDataTransferLength - 1
        
            If ReadData(i) <> Pattern_AU6375(i) Then
              CBWTest_New_128_Sector_AU6375 = 3    'Read fail
           
              Exit Function
            End If
        
        Next
   
        
       CBWTest_New_128_Sector_AU6375 = 1
           
         
    End Function
Public Function CBWTest_New_128_Sector_AU6377(Lun As Byte, PreSlotStatus As Byte) As Byte
Dim i As Long
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long
Dim j As Integer

 CBWDataTransferLength = 65536 ' 64 sector

   
    If PreSlotStatus <> 1 Then
        CBWTest_New_128_Sector_AU6377 = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New_128_Sector_AU6377 = 2
   
    '========================================
    
    
     If OpenPipe = 0 Then
       CBWTest_New_128_Sector_AU6377 = 2   ' Write fail
       Exit Function
     End If
  
    '====================================
     TmpInteger = TestUnitSpeed(Lun)
    
    If TmpInteger = 0 Then
        
       CBWTest_New_128_Sector_AU6377 = 2   ' usb 2.0 high speed fail
       UsbSpeedTestResult = 2
       Exit Function
    End If
    TmpInteger = 0
    
    TmpInteger = TestUnitReady(Lun)
     If TmpInteger = 0 Then
         TmpInteger = RequestSense(Lun)
        
         If TmpInteger = 0 Then
        
            CBWTest_New_128_Sector_AU6377 = 2  'Write fail
            Exit Function
         End If
        
     End If
  
  
   
       
       ' For i = 0 To CBWDataTransferLength - 1
        
       '      ReadData(i) = 0
    
       ' Next

        
     
        TmpInteger = Write_Data_AU6377(LBA, Lun, CBWDataTransferLength)
         
        If TmpInteger = 0 Then
            CBWTest_New_128_Sector_AU6377 = 2  'write fail
         
            Exit Function
        End If
   
        TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
         
        If TmpInteger = 0 Then
            CBWTest_New_128_Sector_AU6377 = 3    'Read fail
             
            Exit Function
        End If
     
        For i = 0 To CBWDataTransferLength - 1
        
            If ReadData(i) <> Pattern_AU6377(i) Then
              CBWTest_New_128_Sector_AU6377 = 3    'Read fail
           
              Exit Function
            End If
        
        Next
   
        
       CBWTest_New_128_Sector_AU6377 = 1
           
         
    End Function

Public Function CBWTest_New_128_Sector_PipeReady(Lun As Byte, PreSlotStatus As Byte) As Byte
Dim i As Long
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long
Dim j As Integer

 CBWDataTransferLength = 65536 ' 64 sector

   
    If PreSlotStatus <> 1 Then
        CBWTest_New_128_Sector_PipeReady = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New_128_Sector_PipeReady = 2
   
    '========================================
    
    
     'If OpenPipe = 0 Then
     '  CBWTest_New_128_Sector_AU6377 = 2   ' Write fail
     '  Exit Function
     'End If
  
    '====================================
    ' TmpInteger = TestUnitSpeed(Lun)
    
    'If TmpInteger = 0 Then
        
    '   CBWTest_New_128_Sector_AU6377 = 2   ' usb 2.0 high speed fail
    '   UsbSpeedTestResult = 2
    '   Exit Function
    'End If
    TmpInteger = 0
    
    'TmpInteger = TestUnitReady(Lun)
    ' If TmpInteger = 0 Then
    '     TmpInteger = RequestSense(Lun)
        
    '     If TmpInteger = 0 Then
        
    '        CBWTest_New_128_Sector_AU6377 = 2  'Write fail
    '        Exit Function
    '     End If
        
    ' End If
  
  
   
       
       ' For i = 0 To CBWDataTransferLength - 1
        
       '      ReadData(i) = 0
    
       ' Next

        
     
        TmpInteger = Write_Data_AU6377(LBA, Lun, CBWDataTransferLength)
         
        If TmpInteger = 0 Then
            CBWTest_New_128_Sector_PipeReady = 2  'write fail
         
            Exit Function
        End If
   
        TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
         
'        If ChipName = "AU6479ULF23" Then
'            If TmpInteger = 0 Then
'                CBWTest_New_128_Sector_PipeReady = 3    'Read fail
'                For i = 0 To 3
'                    Tester.Print "read 64K fail-" & (i + 1)
'                    TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
'                    If TmpInteger <> 0 Then
'                        Exit For
'                    End If
'                Next
'
'                If TmpInteger = 0 And i = 4 Then
'                    Exit Function
'                End If
'            End If
'        Else
            If TmpInteger = 0 Then
                CBWTest_New_128_Sector_PipeReady = 3    'Read fail
                Exit Function
            End If
'        End If
     
        For i = 0 To CBWDataTransferLength - 1
        
            If ReadData(i) <> Pattern_AU6377(i) Then
              CBWTest_New_128_Sector_PipeReady = 3    'Read fail
           
              Exit Function
            End If
        
        Next
   
        
       CBWTest_New_128_Sector_PipeReady = 1
           
         
End Function

Public Function Write_Data_LED(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte

Dim CBW(0 To 30) As Byte
Dim CSW(0 To 12) As Byte
Dim NumberOfBytesWritten As Long
Dim NumberOfBytesRead As Long
Dim CBWDataTransferLen(0 To 3) As Byte
Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim i As Integer
Dim tmpV(0 To 2) As Long
Dim opcode As Byte

opcode = &H2A
'Buffer(0) = &H33 'CByte(Text2.Text)
'Buffer(1) = &H44


    For i = 0 To 30
    
        CBW(i) = 0
    
    Next i
    
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
CBW(12) = &H0                 '80

'////////////// LUN
CBW(13) = Lun                    '00

'///////////// CBD Len
CBW(14) = &HA                '0a

'////////////  UFI command

CBW(15) = opcode
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

 
'1. CBW output
 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)    'out

If result = 0 Then
    Write_Data_LED = 0
    Exit Function
End If
 
 Call MsecDelay(0.2)
CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
 
'2, Output data
result = WriteFile _
       (WriteHandle, _
       Pattern(0), _
       CBWDataTransferLength, _
       NumberOfBytesWritten, _
       0)    'out

 
If result = 0 Then
    Write_Data_LED = 0
    Exit Function
End If

'3 . CSW
result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
        
If result = 0 Then
    Write_Data_LED = 0
    Exit Function
End If
 
 
 
If CSW(12) = 1 Then
Write_Data_LED = 0

Else
Write_Data_LED = 1
End If
End Function


Public Function Write_Data2(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte

Dim CBW(0 To 30) As Byte
Dim CSW(0 To 12) As Byte
Dim NumberOfBytesWritten As Long
Dim NumberOfBytesRead As Long
Dim CBWDataTransferLen(0 To 3) As Byte
Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim i As Integer
Dim tmpV(0 To 2) As Long
Dim opcode As Byte

opcode = &H2A
'Buffer(0) = &H33 'CByte(Text2.Text)
'Buffer(1) = &H44


    For i = 0 To 30
    
        CBW(i) = 0
    
    Next i
    
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
CBW(12) = &H0                 '80

'////////////// LUN
CBW(13) = Lun                    '00

'///////////// CBD Len
CBW(14) = &HA                '0a

'////////////  UFI command

CBW(15) = opcode
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

 
'1. CBW output
 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)    'out

If result = 0 Then
    Write_Data2 = 0
    Exit Function
End If
 
 
 
'2, Output data
result = WriteFile _
       (WriteHandle, _
       Pattern(0), _
       CBWDataTransferLength, _
       NumberOfBytesWritten, _
       0)    'out

 
'If result = 0 Then
 '   Write_Data2 = 0
 '   Exit Function
'End If

'3 . CSW
result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
           
If NumberOfBytesRead <> 13 Then

result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
End If
          
        
If result = 0 Then
    Write_Data2 = 0
    Exit Function
End If
 
 
 
If CSW(12) = 1 Then
Write_Data2 = 0

Else
Write_Data2 = 1
End If
End Function



Public Function Write_Data(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte

Dim CBW(0 To 30) As Byte
Dim CSW(0 To 12) As Byte
Dim NumberOfBytesWritten As Long
Dim NumberOfBytesRead As Long
Dim CBWDataTransferLen(0 To 3) As Byte
Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim i As Integer
Dim tmpV(0 To 2) As Long
Dim opcode As Byte

opcode = &H2A
'Buffer(0) = &H33 'CByte(Text2.Text)
'Buffer(1) = &H44


    For i = 0 To 30
    
        CBW(i) = 0
    
    Next i
    
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
CBW(12) = &H0                 '80

'////////////// LUN
CBW(13) = Lun                    '00

'///////////// CBD Len
CBW(14) = &HA                '0a

'////////////  UFI command

CBW(15) = opcode
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

 
'1. CBW output
 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)    'out

If result = 0 Then
    Write_Data = 0
    Exit Function
End If
 
 
 
'2, Output data
result = WriteFile _
       (WriteHandle, _
       Pattern(0), _
       CBWDataTransferLength, _
       NumberOfBytesWritten, _
       0)    'out

 
If result = 0 Then
    Write_Data = 0
    Exit Function
End If



'3 . CSW
result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in


If result = 0 Then
    Write_Data = 0
    Exit Function
End If
 
 
 
If CSW(12) = 1 Then
Write_Data = 0

Else
Write_Data = 1
End If
End Function
Public Function Write_Data_AssignSize(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte

Dim CBW(0 To 30) As Byte
Dim CSW(0 To 12) As Byte
Dim NumberOfBytesWritten As Long
Dim NumberOfBytesRead As Long
Dim CBWDataTransferLen(0 To 3) As Byte
Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim i As Integer
Dim tmpV(0 To 2) As Long
Dim opcode As Byte

opcode = &H2A
'Buffer(0) = &H33 'CByte(Text2.Text)
'Buffer(1) = &H44


    For i = 0 To 30
    
        CBW(i) = 0
    
    Next i
    
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
CBW(12) = &H0                 '80

'////////////// LUN
CBW(13) = Lun                    '00

'///////////// CBD Len
CBW(14) = &HA                '0a

'////////////  UFI command

CBW(15) = opcode
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

 
'1. CBW output
 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)    'out

If result = 0 Then
    Write_Data_AssignSize = 0
    Exit Function
End If
 
If (result = 0) Or (NumberOfBytesWritten <> 31) Then
    Write_Data_AssignSize = 0
    Exit Function
End If

'2, Output data
result = WriteFile _
       (WriteHandle, _
       Pattern(0), _
       CBWDataTransferLength, _
       NumberOfBytesWritten, _
       0)    'out

 
If (result = 0) Or (NumberOfBytesWritten <> CBWDataTransferLength) Then
    Write_Data_AssignSize = 0
    Exit Function
End If

'3 . CSW
result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in


If result = 0 Then
    Write_Data_AssignSize = 0
    Exit Function
End If
 
 
 
If CSW(12) = 1 Then
Write_Data_AssignSize = 0

Else
Write_Data_AssignSize = 1
End If

End Function

Public Function Write_EEPromData(AddressHighByte As Byte, AddressLowByte As Byte, EEPromData As Byte) As Byte

Dim CBW(0 To 30) As Byte
Dim CSW(0 To 12) As Byte
Dim NumberOfBytesWritten As Long
Dim NumberOfBytesRead As Long
Dim CBWDataTransferLen(0 To 3) As Byte
Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim i As Integer
Dim tmpV(0 To 2) As Long
Dim opcode As Byte

opcode = &HC1
'Buffer(0) = &H33 'CByte(Text2.Text)
'Buffer(1) = &H44


    For i = 0 To 30
    
        CBW(i) = 0
    
    Next i
    
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

 
CBW(8) = 0  '00
CBW(9) = 0  '08
CBW(10) = 0 '00
CBW(11) = 0 '00

'///////////////  CBW Flag
CBW(12) = &H0                 '80

'////////////// LUN
CBW(13) = 0                    '00

'///////////// CBD Len
CBW(14) = &HA                '0a

'////////////  UFI command

CBW(15) = opcode
CBW(16) = AddressHighByte
 
  

CBW(17) = AddressLowByte
CBW(18) = EEPromData
 

For i = 19 To 30
    CBW(i) = 0
Next

 
'1. CBW output
 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)    'out

If result = 0 Then
    Write_EEPromData = 0
    Exit Function
End If
 
 
 
'3 . CSW
result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
        
If result = 0 Then
    Write_EEPromData = 0
    Exit Function
End If
 
 
 
If CSW(12) = 1 Then
Write_EEPromData = 0

Else
Write_EEPromData = 1
End If
End Function
Public Function Write_Data_WPTest(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte

Dim CBW(0 To 30) As Byte
Dim CSW(0 To 12) As Byte
Dim NumberOfBytesWritten As Long
Dim NumberOfBytesRead As Long
Dim CBWDataTransferLen(0 To 3) As Byte
Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim i As Integer
Dim tmpV(0 To 2) As Long
Dim opcode As Byte

opcode = &H2A
'Buffer(0) = &H33 'CByte(Text2.Text)
'Buffer(1) = &H44


    For i = 0 To 30
    
        CBW(i) = 0
    
    Next i
    
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
CBW(12) = &H0                 '80

'////////////// LUN
CBW(13) = Lun                    '00

'///////////// CBD Len
CBW(14) = &HA                '0a

'////////////  UFI command

CBW(15) = opcode
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

 
'1. CBW output
 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)    'out

If result = 0 Then
    Write_Data_WPTest = 0
    Exit Function
End If
 
 
 
'2, Output data
result = WriteFile _
       (WriteHandle, _
       Pattern(0), _
       CBWDataTransferLength, _
       NumberOfBytesWritten, _
       0)    'out

 
 

'3 . CSW
result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
        
If result = 0 Then
    result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
    
End If
 
 
 
 
If CSW(12) = 1 Then
   result = RequestSense(Lun)
   If RequestSenseData(12) = 39 Then
    
       Write_Data_WPTest = 1 'Write fail
       Else
       Write_Data_WPTest = 2
       
    End If
    Exit Function
Else
Write_Data_WPTest = 2
End If
End Function

Public Function Read_Data(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte
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
 Read_Data = 0
 Exit Function
End If

'2. Readdata stage
 
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in


 
'If result = 0 Then
' Read_Data = 0
' Exit Function
'End If

'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
If result = 0 Then
 Read_Data = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
    Read_Data = 0
Else
     Read_Data = 1
   
End If

 
End Function
Public Function Read_6990HW_Code(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte
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
CBW(13) = &H0                    '00

'///////////// CBD Len
CBW(14) = &HA                '0a

'////////////  UFI command

CBW(15) = &HFA
CBW(16) = &HE
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
 Read_6990HW_Code = 0
 Exit Function
End If

'2. Readdata stage
 
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in


 
'If result = 0 Then
' Read_Data = 0
' Exit Function
'End If

'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
If result = 0 Then
 Read_6990HW_Code = 0
 Exit Function
End If
 
'4. CSW status

If ReadData(11) = 0 Then        'version "A" or "R"
    Read_6990HW_Code = 1
ElseIf ReadData(11) = 32 Then   'version "B"
    Read_6990HW_Code = 2
ElseIf ReadData(11) = 64 Then   'version "S"
    Read_6990HW_Code = 3
Else
    Read_6990HW_Code = 0
End If

 
End Function

Public Function Read_6922HW_Code(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte
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
CBW(13) = &H0                    '00

'///////////// CBD Len
CBW(14) = &HA                '0a

'////////////  UFI command

CBW(15) = &HFA
CBW(16) = &HE
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
 Read_6922HW_Code = 0
 Exit Function
End If

'2. Readdata stage
 
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in


 
'If result = 0 Then
' Read_Data = 0
' Exit Function
'End If

'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
If result = 0 Then
 Read_6922HW_Code = 0
 Exit Function
End If
 
'4. CSW status

If ReadData(11) = &H80 Then       'H62
    Read_6922HW_Code = 1
ElseIf ReadData(11) = &HC0 Then   'I62
    Read_6922HW_Code = 2
Else
    Read_6922HW_Code = 0
End If

 
End Function


Public Function Read_Data2(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte
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
 Read_Data2 = 0
 Exit Function
End If

'2. Readdata stage
 Call MsecDelay(1.2)
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

 
 

'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
 If NumberOfBytesRead <> 13 Then
 
  result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 End If
 
 
 
If result = 0 Then
 Read_Data2 = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
    Read_Data2 = 0
Else
     Read_Data2 = 1
   
End If

 
End Function
Public Function Read_DataCIS(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte
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

 

'2. Readdata stage
  
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
          
 If NumberOfBytesRead = 2048 Then
 
    Read_DataCIS = 1
    Exit Function
 End If

'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
 
 
'4. CSW status

 

 
End Function

 Public Function Read_Data1(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte
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
 Read_Data1 = 0
 Exit Function
End If

'2. Readdata stage
  
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

 
 

'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
 If NumberOfBytesRead <> 13 Then
 
  result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 End If
 
 
 
If result = 0 Then
 Read_Data1 = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
    Read_Data1 = 0
Else
     Read_Data1 = 1
   
End If

 
End Function
Public Function Read_Data_AssignSize(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte
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
 Read_Data_AssignSize = 0
 Exit Function
End If

'2. Readdata stage

result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
          CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in



If (result = 0) Or (NumberOfBytesRead <> CBWDataTransferLength) Then
    Read_Data_AssignSize = 0
    Exit Function
End If
 

'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
 If NumberOfBytesRead <> 13 Then
 
  result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 End If
 
 
 
If result = 0 Then
 Read_Data_AssignSize = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
    Read_Data_AssignSize = 0
Else
     Read_Data_AssignSize = 1
   
End If

 
End Function

Public Function Read_SD_Speed_AU6371(LBA As Long, Lun As Byte, CBWDataTransferLength As Long, BitWidth As String) As Byte
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

'////////////  ve

CBW(15) = &HC4
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
 Read_SD_Speed_AU6371 = 0
 Exit Function
End If

'2. Readdata stage
  
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

' For i = 0 To 15
' Debug.Print "i="; i; Hex(ReadData(i))
' Next

 

 

Read_SD_Speed_AU6371 = 0

If BitWidth = "8Bits" Then
 Tester.Print "SD bus width="; Hex(ReadData(14))
  If (ReadData(14) = &H78) Or (ReadData(14) = &H68) Then
      Read_SD_Speed_AU6371 = 1
      Tester.Print "SD bus width is 8 bits, 78 is48 MHZ, 68 is 24 MHZ"
  End If
  
End If
  
  
If BitWidth = "4Bits" Then
   Tester.Print "SD bus width="; Hex(ReadData(14))
  If (ReadData(14) = &HF0) Then
      Read_SD_Speed_AU6371 = 1
      Tester.Print "SD bus width is 4 bits,48 MHZ"
  End If
  
End If
    
    
If Read_SD_Speed_AU6371 = 0 Then
Exit Function
End If
  


'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
 
 
 
 If NumberOfBytesRead <> 13 Then
 
  result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 End If
 
 
 
If result = 0 Then
 Read_SD_Speed_AU6371 = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
     Read_SD_Speed_AU6371 = 0
Else
     Read_SD_Speed_AU6371 = 1
   
End If

 
End Function

Public Function Read_SD_Speed(LBA As Long, Lun As Byte, CBWDataTransferLength As Long, BitWidth As String) As Byte
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

'////////////  ve

CBW(15) = &HC4
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
 Read_SD_Speed = 0
 Exit Function
End If

'2. Readdata stage
  
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

' For i = 0 To 15
' Debug.Print "i="; i; Hex(ReadData(i))
' Next

 

 

Read_SD_Speed = 0

If BitWidth = "8Bits" Then
 Tester.Print "SD bus width="; Hex(ReadData(14))
  If (ReadData(14) = &H78) Then
      Read_SD_Speed = 1
      Tester.Print "SD bus width is 8 bits, 48 MHZ"
  End If
  
End If
  
  
If BitWidth = "4Bits" Then
   Tester.Print "SD bus width="; Hex(ReadData(14))
  If (ReadData(14) = &HF0) Then
      Read_SD_Speed = 1
      Tester.Print "SD bus width is 4 bits,48 MHZ"
  End If
  
End If
    
    
If Read_SD_Speed = 0 Then
Exit Function
End If
  


'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
 
 
 
 If NumberOfBytesRead <> 13 Then
 
  result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 End If
 
 
 
If result = 0 Then
 Read_SD_Speed = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
     Read_SD_Speed = 0
Else
     Read_SD_Speed = 1
   
End If

 
End Function

Public Function Read_Speed2ReadData(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte
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

For i = 0 To CBWDataTransferLength - 1
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



'////////////  CDB
CBW(15) = &HC4
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
    Read_Speed2ReadData = 0
    Exit Function
End If

'2. Readdata stage
  
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
          CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

If result = 0 Then
    Read_Speed2ReadData = 0
    Exit Function
End If

Read_Speed2ReadData = 0

'3. CSW data
result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
 
If NumberOfBytesRead <> 13 Then
 
    result = ReadFile _
            (ReadHandle, _
             CSW(0), _
             13, _
             NumberOfBytesRead, _
             HIDOverlapped)  'in
End If
 
 
 
If result = 0 Then
    Read_Speed2ReadData = 0
    Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
     Read_Speed2ReadData = 0
Else
     Read_Speed2ReadData = 1
   
End If
 
End Function

Public Function Read_SD_Speed_AU6476_48MHz(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte
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

'////////////  ve

CBW(15) = &HC4
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
 Read_SD_Speed_AU6476_48MHz = 0
 Exit Function
End If

'2. Readdata stage
  
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

' For i = 0 To 15
' Debug.Print "i="; i; Hex(ReadData(i))
' Next

 

 

Read_SD_Speed_AU6476_48MHz = 0


Tester.Print "SD bus width="; Hex(ReadData(14))
If (ReadData(14) = &H78) Then
    Read_SD_Speed_AU6476_48MHz = 1
    Tester.Print "SD bus width is 8 bits, 48 MHZ"
End If
  
If (ReadData(14) = &HF0) Then
    Read_SD_Speed_AU6476_48MHz = 1
    Tester.Print "SD bus width is 4 bits,48 MHZ"
End If
    
    
If Read_SD_Speed_AU6476_48MHz = 0 Then
    Exit Function
End If
  


'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
 
 
 
 If NumberOfBytesRead <> 13 Then
 
  result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 End If
 
 
 
If result = 0 Then
 Read_SD_Speed_AU6476_48MHz = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
     Read_SD_Speed_AU6476_48MHz = 0
Else
     Read_SD_Speed_AU6476_48MHz = 1
   
End If

 
End Function
Public Function Read_SD_Speed_AU6435(LBA As Long, Lun As Byte, CBWDataTransferLength As Long, BitWidth As String) As Byte
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


Const CBWTag_0 = &H18
Const CBWTag_1 = &H9A
Const CBWTag_2 = &H20
Const CBWTag_3 = &H88


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

'CBW(8) = CBWDataTransferLen(0)  '00
'CBW(9) = CBWDataTransferLen(1)  '08
'CBW(10) = CBWDataTransferLen(2) '00
'CBW(11) = CBWDataTransferLen(3) '00
CBW(8) = &H40
CBW(9) = &H0
CBW(10) = &H0
CBW(11) = &H0


'///////////////  CBW Flag
CBW(12) = &H80                 '80

'////////////// LUN
CBW(13) = Lun                    '00

'///////////// CBD Len
CBW(14) = &H10                 '0a

'////////////  ve

CBW(15) = &HC7
'CBW(16) = Lun * 32

CBW(16) = &H1F

'LBAByte(0) = (LBA Mod 256)
'tmpV(0) = Int(LBA / 256)
'LBAByte(1) = (tmpV(0) Mod 256)
'tmpV(1) = Int(tmpV(0) / 256)
'LBAByte(2) = (tmpV(1) Mod 256)
'tmpV(2) = Int((tmpV(1) / 256))
'LBAByte(3) = (tmpV(2) Mod 256)

'CBW(17) = LBAByte(3)         '00
'CBW(18) = LBAByte(2)         '00
'CBW(19) = LBAByte(1)         '00
'CBW(20) = LBAByte(0)         '40

CBW(17) = &H5
CBW(18) = &H8F
CBW(19) = &HC4
CBW(20) = &H0



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
 Read_SD_Speed_AU6435 = 0
 Exit Function
End If

'2. Readdata stage
  
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

' For i = 0 To 15
' Debug.Print "i="; i; Hex(ReadData(i))
' Next

 

 

Read_SD_Speed_AU6435 = 0

If BitWidth = "8Bits" Then
 Tester.Print "SD bus width="; Hex(ReadData(15))
  If (ReadData(15) = &H72) Then
      Read_SD_Speed_AU6435 = 1
      Tester.Print "SD bus width is 8 bits, 48 MHZ"
  End If
  
End If
  
  
If BitWidth = "4Bits" Then
   Tester.Print "SD bus width="; Hex(ReadData(15))
  If (ReadData(15) = &H71) Then
      Read_SD_Speed_AU6435 = 1
      Tester.Print "SD bus width is 4 bits,48 MHZ"
  End If
  
End If
    
    
If Read_SD_Speed_AU6435 = 0 Then
Exit Function
End If
  


'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
 
 
 
 If NumberOfBytesRead <> 13 Then
 
  result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 End If
 
 
 
If result = 0 Then
 Read_SD_Speed_AU6435 = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
     Read_SD_Speed_AU6435 = 0
Else
     Read_SD_Speed_AU6435 = 1
   
End If

 
End Function
Public Function Read_SD30_Speed_AU6435(LBA As Long, Lun As Byte, CBWDataTransferLength As Long, BitWidth As String) As Byte
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


Const CBWTag_0 = &H18
Const CBWTag_1 = &H9A
Const CBWTag_2 = &H20
Const CBWTag_3 = &H88


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

'CBW(8) = CBWDataTransferLen(0)  '00
'CBW(9) = CBWDataTransferLen(1)  '08
'CBW(10) = CBWDataTransferLen(2) '00
'CBW(11) = CBWDataTransferLen(3) '00
CBW(8) = &H40
CBW(9) = &H0
CBW(10) = &H0
CBW(11) = &H0


'///////////////  CBW Flag
CBW(12) = &H80                 '80

'////////////// LUN
CBW(13) = Lun                    '00

'///////////// CBD Len
CBW(14) = &H10                 '0a

'////////////  ve

CBW(15) = &HC7
'CBW(16) = Lun * 32

CBW(16) = &H1F

'LBAByte(0) = (LBA Mod 256)
'tmpV(0) = Int(LBA / 256)
'LBAByte(1) = (tmpV(0) Mod 256)
'tmpV(1) = Int(tmpV(0) / 256)
'LBAByte(2) = (tmpV(1) Mod 256)
'tmpV(2) = Int((tmpV(1) / 256))
'LBAByte(3) = (tmpV(2) Mod 256)

'CBW(17) = LBAByte(3)         '00
'CBW(18) = LBAByte(2)         '00
'CBW(19) = LBAByte(1)         '00
'CBW(20) = LBAByte(0)         '40

CBW(17) = &H5
CBW(18) = &H8F
CBW(19) = &HC4
CBW(20) = &H0



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
 Read_SD30_Speed_AU6435 = 0
 Exit Function
End If

'2. Readdata stage
  
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

' For i = 0 To 15
' Debug.Print "i="; i; Hex(ReadData(i))
' Next

 

 

Read_SD30_Speed_AU6435 = 0

If BitWidth = "8Bits" Then
 Tester.Print "SD bus width="; Hex(ReadData(15))
  If (ReadData(15) = &H82) Then
      Read_SD30_Speed_AU6435 = 1
      Tester.Print "SD bus width is 8 bits, 60 MHZ"
  End If
  
End If
  
  
If BitWidth = "4Bits" Then
   Tester.Print "SD bus width="; Hex(ReadData(15))
  If (ReadData(15) = &H81) Then
      Read_SD30_Speed_AU6435 = 1
      Tester.Print "SD bus width is 4 bits,60 MHZ"
  End If
  
  If (ReadData(15) = &H91) Then
      Read_SD30_Speed_AU6435 = 1
      Tester.Print "SD bus width is 4 bits,100 MHZ"
  End If
  
  If (ReadData(15) = &H99) Then
      Read_SD30_Speed_AU6435 = 1
      Tester.Print "SD bus width is 4 bits,100 MHZ"
  End If
  
End If
    
    
If Read_SD30_Speed_AU6435 = 0 Then
Exit Function
End If
  


'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
 
 
 
 If NumberOfBytesRead <> 13 Then
 
  result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 End If
 
 
 
If result = 0 Then
 Read_SD30_Speed_AU6435 = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
     Read_SD30_Speed_AU6435 = 0
Else
     Read_SD30_Speed_AU6435 = 1
   
End If

 
End Function
Public Function Read_SD30_Speed_AU6435_100(LBA As Long, Lun As Byte, CBWDataTransferLength As Long, BitWidth As String) As Byte
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


Const CBWTag_0 = &H18
Const CBWTag_1 = &H9A
Const CBWTag_2 = &H20
Const CBWTag_3 = &H88


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

'CBW(8) = CBWDataTransferLen(0)  '00
'CBW(9) = CBWDataTransferLen(1)  '08
'CBW(10) = CBWDataTransferLen(2) '00
'CBW(11) = CBWDataTransferLen(3) '00
CBW(8) = &H40
CBW(9) = &H0
CBW(10) = &H0
CBW(11) = &H0


'///////////////  CBW Flag
CBW(12) = &H80                 '80

'////////////// LUN
CBW(13) = Lun                    '00

'///////////// CBD Len
CBW(14) = &H10                 '0a

'////////////  ve

CBW(15) = &HC7
'CBW(16) = Lun * 32

CBW(16) = &H1F

'LBAByte(0) = (LBA Mod 256)
'tmpV(0) = Int(LBA / 256)
'LBAByte(1) = (tmpV(0) Mod 256)
'tmpV(1) = Int(tmpV(0) / 256)
'LBAByte(2) = (tmpV(1) Mod 256)
'tmpV(2) = Int((tmpV(1) / 256))
'LBAByte(3) = (tmpV(2) Mod 256)

'CBW(17) = LBAByte(3)         '00
'CBW(18) = LBAByte(2)         '00
'CBW(19) = LBAByte(1)         '00
'CBW(20) = LBAByte(0)         '40

CBW(17) = &H5
CBW(18) = &H8F
CBW(19) = &HC4
CBW(20) = &H0



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
 Read_SD30_Speed_AU6435_100 = 0
 Exit Function
End If

'2. Readdata stage
  
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

' For i = 0 To 15
' Debug.Print "i="; i; Hex(ReadData(i))
' Next

 

 

Read_SD30_Speed_AU6435_100 = 0

If BitWidth = "8Bits" Then
 Tester.Print "SD bus width="; Hex(ReadData(15))
  If (ReadData(15) = &H82) Then
      Tester.Print "SD bus width is 8 bits, 60 MHZ"
  End If
  
End If
  
  
If BitWidth = "4Bits" Then
   Tester.Print "SD bus width="; Hex(ReadData(15))
  If (ReadData(15) = &H81) Then
      Tester.Print "SD bus width is 4 bits,60 MHZ"
  End If
  
  If (ReadData(15) = &H91) Then
      Read_SD30_Speed_AU6435_100 = 1
      Tester.Print "SD bus width is 4 bits,100 MHZ"
  End If
  
End If
    
    
If Read_SD30_Speed_AU6435_100 = 0 Then
Exit Function
End If
  


'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
 
 
 
 If NumberOfBytesRead <> 13 Then
 
  result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 End If
 
 
 
If result = 0 Then
 Read_SD30_Speed_AU6435_100 = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
     Read_SD30_Speed_AU6435_100 = 0
Else
     Read_SD30_Speed_AU6435_100 = 1
   
End If

 
End Function
Public Function Read_SD30_Mode_AU6435(LBA As Long, Lun As Byte, CBWDataTransferLength As Long, AccessMode As String) As Byte
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
Dim TmpValue As Byte

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


Const CBWTag_0 = &H10
Const CBWTag_1 = &HF4
Const CBWTag_2 = &H2C
Const CBWTag_3 = &H89


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

'CBW(8) = CBWDataTransferLen(0)  '00
'CBW(9) = CBWDataTransferLen(1)  '08
'CBW(10) = CBWDataTransferLen(2) '00
'CBW(11) = CBWDataTransferLen(3) '00
CBW(8) = &H8
CBW(9) = &H0
CBW(10) = &H0
CBW(11) = &H0


'///////////////  CBW Flag
CBW(12) = &H80                 '80

'////////////// LUN
CBW(13) = Lun                    '00

'///////////// CBD Len
CBW(14) = &H10                 '0a

'////////////  ve

CBW(15) = &HC7
'CBW(16) = Lun * 32

CBW(16) = &H1F

'LBAByte(0) = (LBA Mod 256)
'tmpV(0) = Int(LBA / 256)
'LBAByte(1) = (tmpV(0) Mod 256)
'tmpV(1) = Int(tmpV(0) / 256)
'LBAByte(2) = (tmpV(1) Mod 256)
'tmpV(2) = Int((tmpV(1) / 256))
'LBAByte(3) = (tmpV(2) Mod 256)

'CBW(17) = LBAByte(3)         '00
'CBW(18) = LBAByte(2)         '00
'CBW(19) = LBAByte(1)         '00
'CBW(20) = LBAByte(0)         '40

CBW(17) = &H5
CBW(18) = &H8F
CBW(19) = &HC7
CBW(20) = &H84



'/////////////  Reverve
CBW(21) = &H30


'//////////// Transfer Len

'TransferLen = Int(CBWDataTransferLength / 512)

'TransferLenLSB = (TransferLen Mod 256)
'tmpV(0) = Int(TransferLen / 256)
'TransferLenMSB = (tmpV(0) / 256)

CBW(22) = &H35
CBW(23) = &H38
CBW(24) = &H46

For i = 25 To 30
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
 Read_SD30_Mode_AU6435 = 0
 Exit Function
End If

'2. Readdata stage
  
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

' For i = 0 To 15
' Debug.Print "i="; i; Hex(ReadData(i))
' Next

 

 

Read_SD30_Mode_AU6435 = 0

  
'TmpValue = CAndValue(ReadData(2), &HF0)
    

If AccessMode = "Non-UHS" Then
    Tester.Print "SD Mode ="; Hex(ReadData(0))
    
    If (ReadData(0) = &H1) Then
        Read_SD30_Mode_AU6435 = 1
        Tester.Print "SD Mode is Non-UHS"
    End If

ElseIf AccessMode = "SDR" Then
    
    Tester.Print "SD Mode ="; Hex(ReadData(0))
    If (ReadData(0) = &H2) Then
        Read_SD30_Mode_AU6435 = 1
        Tester.Print "SD Mode is SDR"
    End If
    
    If (ReadData(0) = &H3) Then
        Read_SD30_Mode_AU6435 = 1
        Tester.Print "SD Mode is DDR"
    End If
    

ElseIf AccessMode = "DDR" Then
    
    Tester.Print "SD Mode ="; Hex(ReadData(0))
    If (ReadData(0) = &H3) Then
        Read_SD30_Mode_AU6435 = 1
        Tester.Print "SD Mode is DDR"
    End If
    
End If
    
    
If Read_SD30_Mode_AU6435 = 0 Then
Exit Function
End If
  


'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
 
 
 
 If NumberOfBytesRead <> 13 Then
 
  result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 End If
 
 
 
If result = 0 Then
 Read_SD30_Mode_AU6435 = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
     Read_SD30_Mode_AU6435 = 0
Else
     Read_SD30_Mode_AU6435 = 1
   
End If

 
End Function
Public Function Read_MS_Speed_AU6435(LBA As Long, Lun As Byte, CBWDataTransferLength As Long, BitWidth As String) As Byte
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
Dim TmpValue As Byte

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


Const CBWTag_0 = &H18
Const CBWTag_1 = &H9A
Const CBWTag_2 = &H20
Const CBWTag_3 = &H88


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

'CBW(8) = CBWDataTransferLen(0)  '00
'CBW(9) = CBWDataTransferLen(1)  '08
'CBW(10) = CBWDataTransferLen(2) '00
'CBW(11) = CBWDataTransferLen(3) '00
CBW(8) = &H40
CBW(9) = &H0
CBW(10) = &H0
CBW(11) = &H0


'///////////////  CBW Flag
CBW(12) = &H80                 '80

'////////////// LUN
CBW(13) = Lun                    '00

'///////////// CBD Len
CBW(14) = &H10                 '0a

'////////////  ve

CBW(15) = &HC7
'CBW(16) = Lun * 32

CBW(16) = &H1F

'LBAByte(0) = (LBA Mod 256)
'tmpV(0) = Int(LBA / 256)
'LBAByte(1) = (tmpV(0) Mod 256)
'tmpV(1) = Int(tmpV(0) / 256)
'LBAByte(2) = (tmpV(1) Mod 256)
'tmpV(2) = Int((tmpV(1) / 256))
'LBAByte(3) = (tmpV(2) Mod 256)

'CBW(17) = LBAByte(3)         '00
'CBW(18) = LBAByte(2)         '00
'CBW(19) = LBAByte(1)         '00
'CBW(20) = LBAByte(0)         '40

CBW(17) = &H5
CBW(18) = &H8F
CBW(19) = &HC4
CBW(20) = &H0



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
 Read_MS_Speed_AU6435 = 0
 Exit Function
End If

'2. Readdata stage
  
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

' For i = 0 To 15
' Debug.Print "i="; i; Hex(ReadData(i))
' Next

 

 

Read_MS_Speed_AU6435 = 0

If BitWidth = "8Bits" Then
 TmpValue = CAndValue(ReadData(17), &HF0)   'just compare H-byte
 Tester.Print "MS bus width="; TmpValue
  If (TmpValue = &H90) Then
      Read_MS_Speed_AU6435 = 1
      Tester.Print "MS bus width is 8 bits, 40 MHZ"
  End If
  
End If
  
  
If BitWidth = "4Bits" Then
    TmpValue = CAndValue(ReadData(17), &HF0)   'just compare H-byte
    Tester.Print "MS bus width="; TmpValue
  If (TmpValue = &H50) Then
      Read_MS_Speed_AU6435 = 1
      Tester.Print "MS bus width is 4 bits,40 MHZ"
  End If
  
End If
    
    
If Read_MS_Speed_AU6435 = 0 Then
Exit Function
End If
  


'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
 
 
 
 If NumberOfBytesRead <> 13 Then
 
  result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 End If
 
 
 
If result = 0 Then
 Read_MS_Speed_AU6435 = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
     Read_MS_Speed_AU6435 = 0
Else
     Read_MS_Speed_AU6435 = 1
   
End If

 
End Function
Public Function Read_SD_SpeedE55(LBA As Long, Lun As Byte, CBWDataTransferLength As Long, BitWidth As String) As Byte
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
CBW(13) = &H0
CBW(14) = &H10
CBW(15) = &HC7
CBW(16) = &H1F
CBW(17) = &H5
CBW(18) = &H8F

'////////////// LUN
'CBW(13) = Lun                   '00

'///////////// CBD Len
'CBW(14) = &HA                '0a

'////////////  ve

CBW(19) = &HC4
 CBW(20) = Lun * 32
'LBAByte(0) = (LBA Mod 256)
'tmpV(0) = Int(LBA / 256)
'LBAByte(1) = (tmpV(0) Mod 256)
'tmpV(1) = Int(tmpV(0) / 256)
'LBAByte(2) = (tmpV(1) Mod 256)
'tmpV(2) = Int((tmpV(1) / 256))
'LBAByte(3) = (tmpV(2) Mod 256)

 CBW(21) = LBAByte(3)         '00
 CBW(22) = LBAByte(2)         '00
 CBW(23) = LBAByte(1)         '00
 CBW(24) = LBAByte(0)         '40

'/////////////  Reverve
 CBW(25) = 0

'//////////// Transfer Len

'TransferLen = Int(CBWDataTransferLength / 512)

'TransferLenLSB = (TransferLen Mod 256)
'tmpV(0) = Int(TransferLen / 256)
'TransferLenMSB = (tmpV(0) / 256)

 CBW(26) = TransferLenMSB      '00
 CBW(27) = TransferLenLSB      '04

 For i = 28 To 30
     CBW(i) = 0
 Next

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
 
Dim result As Long

'1. CBW command

result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in  clear stall
 
 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)    'out

If result = 0 Then
 Read_SD_SpeedE55 = 0
 Exit Function
End If

'2. Readdata stage
  
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

' For i = 0 To 15
' Debug.Print "i="; i; Hex(ReadData(i))
' Next

 

 

Read_SD_SpeedE55 = 0

If BitWidth = "8Bits" Then
 Tester.Print "SD bus width="; Hex(ReadData(14))
  If (ReadData(14) = &H78) Then
      Read_SD_SpeedE55 = 1
      Tester.Print "SD bus width is 8 bits, 48 MHZ"
  End If
  
End If
  
  
If BitWidth = "4Bits" Then
   Tester.Print "SD bus width="; Hex(ReadData(14))
  If (ReadData(14) = &HF0) Then
      Read_SD_SpeedE55 = 1
      Tester.Print "SD bus width is 4 bits,48 MHZ"
  End If
  
End If
    
    
If Read_SD_SpeedE55 = 0 Then
Exit Function
End If
  


'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
 
 
 
 If NumberOfBytesRead <> 13 Then
 
  result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 End If
 
 
 
If result = 0 Then
 Read_SD_SpeedE55 = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
     Read_SD_SpeedE55 = 0
Else
     Read_SD_SpeedE55 = 1
   
End If

 
End Function
Public Function Read_SD_Speed_AU6473(LBA As Long, Lun As Byte, CBWDataTransferLength As Long, BitWidth As String) As Byte
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

'////////////  ve

CBW(15) = &HC4
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
 Read_SD_Speed_AU6473 = 0
 Exit Function
End If

'2. Readdata stage
  
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

' For i = 0 To 15
' Debug.Print "i="; i; Hex(ReadData(i))
' Next

Read_SD_Speed_AU6473 = 0

If Lun = 0 Then

    If BitWidth = "8Bits" Then
        Tester.Print "SD bus width="; Hex(ReadData(15))
        If (ReadData(15) = &H78) Then
            Read_SD_Speed_AU6473 = 1
            Tester.Print "SD bus width is 8 bits, 48 MHZ"
        End If
    End If
  
  
    If BitWidth = "4Bits" Then
        Tester.Print "SD bus width="; Hex(ReadData(15))
        If (ReadData(15) = &H71) Then
            Read_SD_Speed_AU6473 = 1
            Tester.Print "SD bus width is 4 bits,48 MHZ"
        End If
  
    End If
    
Else 'Lun1
    
    If BitWidth = "8Bits" Then
        Tester.Print "SD bus width="; Hex(ReadData(22))
        If (ReadData(22) = &H78) Then
            Read_SD_Speed_AU6473 = 1
            Tester.Print "SD bus width is 8 bits, 48 MHZ"
        End If
    End If
  
  
    If BitWidth = "4Bits" Then
        Tester.Print "SD bus width="; Hex(ReadData(22))
        If (ReadData(22) = &H71) Then
            Read_SD_Speed_AU6473 = 1
            Tester.Print "SD bus width is 4 bits,48 MHZ"
        End If
  
    End If
    
End If
    
    
If Read_SD_Speed_AU6473 = 0 Then
Exit Function
End If
  


'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
 
 
 
 If NumberOfBytesRead <> 13 Then
 
  result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 End If
 
 
 
If result = 0 Then
 Read_SD_Speed_AU6473 = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
     Read_SD_Speed_AU6473 = 0
Else
     Read_SD_Speed_AU6473 = 1
   
End If

 
End Function
Public Function Read_OverCurrent5(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte
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

'////////////  ve

CBW(15) = &HC4
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
 Read_OverCurrent5 = 0
 Exit Function
End If

'2. Readdata stage
  
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

  For i = 0 To 15
  Debug.Print "i="; i; Hex(ReadData(i))
  Next

 

 

Read_OverCurrent5 = 0

 
 Tester.Print "OverCurrent Status"; Hex(ReadData(4))
  If (ReadData(4) = &H0) Then
      Read_OverCurrent5 = 1
      Tester.Print "Read_OverCurrent5 ok"
  Else
       Read_OverCurrent5 = 3 'true overcurrent
       Tester.Print "true Read_OverCurrent5 fail"
  End If
  
 
  
      
    
If Read_OverCurrent5 <> 1 Then
Exit Function
End If
  


'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
 
 
 
 If NumberOfBytesRead <> 13 Then
 
  result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 End If
 
 
 
If result = 0 Then
 Read_OverCurrent5 = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
     Read_OverCurrent5 = 0
Else
     Read_OverCurrent5 = 1
   
End If

 
End Function

Public Function Read_OverCurrent(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte
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

'////////////  ve

CBW(15) = &HC4
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
 Read_OverCurrent = 0
 Exit Function
End If

'2. Readdata stage
  
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

  For i = 0 To 15
  Debug.Print "i="; i; Hex(ReadData(i))
  Next

 

 

Read_OverCurrent = 0

 
 Tester.Print "OverCurrent Status"; Hex(ReadData(4))
  If (ReadData(4) = &H0) Then
      Read_OverCurrent = 1
      Tester.Print "Read_OverCurrent ok"
  Else
       Read_OverCurrent = 3 'true overcurrent
       Tester.Print "true Read_OverCurrent fail"
  End If
  
 
  
      
    
If Read_OverCurrent <> 1 Then
Exit Function
End If
  


'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
 
 
 
 If NumberOfBytesRead <> 13 Then
 
  result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 End If
 
 
 
If result = 0 Then
 Read_OverCurrent = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
     Read_OverCurrent = 0
Else
     Read_OverCurrent = 1
   
End If

 
End Function


Public Function Read_EEPRomData(AddressHighByte As Byte, AddressLowByte As Byte, CBWDataTransferLength As Long) As Byte
Dim CBW(0 To 30) As Byte
Dim NumberOfBytesWritten As Long
Dim CBWDataTransferLen(0 To 3) As Byte
  
Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim i As Long
Dim tmpV(0 To 2) As Long
Dim opcode As Byte
opcode = &HC0
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

'////////////  ve

CBW(15) = opcode
CBW(16) = AddressHighByte
CBW(17) = AddressLowByte

 
For i = 18 To 30
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
 Read_EEPRomData = 0
 Exit Function
End If

'2. Readdata stage
  
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

 For i = 0 To 6
  Tester.Print i; Hex(ReadData(i));
  
Next i
 

 Tester.Print

 
 
 
  
      
    
   


'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
 
 
 
 If NumberOfBytesRead <> 13 Then
 
  result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 End If
 
 
 
If result = 0 Then
 Read_EEPRomData = 0
 Exit Function
End If
 
'4. CSW status

 
     Read_EEPRomData = 1
      For i = 0 To 6
          If ReadData(i) <> i Then
                Read_EEPRomData = 0
          Exit For
          End If
      Next
      
 

 
End Function
Public Function Read_GPI_AU6476(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte
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

'////////////  ve

CBW(15) = &H77
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
 Read_GPI_AU6476 = 0
 Exit Function
End If

'2. Readdata stage
  
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

 For i = 0 To 3
  Debug.Print "i="; i; Hex(ReadData(i))
  
Next i
 

 

Read_GPI_AU6476 = ReadData(2)

 
 
  
      
    
   


'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
 
 
 
 If NumberOfBytesRead <> 13 Then
 
  result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 End If
 
 
 
If result = 0 Then
 Read_GPI_AU6476 = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
     Read_GPI_AU6476 = 0
 End If

 
End Function
Public Function Read_GPI(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte
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

'////////////  ve

CBW(15) = &HC4
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
 Read_GPI = 0
 Exit Function
End If

'2. Readdata stage
  
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

 For i = 0 To 30
  Debug.Print "i="; i; Hex(ReadData(i))
  
Next i
 

 

Read_GPI = ReadData(30)

 
 
  
      
    
   


'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
 
 
 
 If NumberOfBytesRead <> 13 Then
 
  result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 End If
 
 
 
If result = 0 Then
 Read_GPI = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
     Read_GPI = 0
 End If

 
End Function

Public Function AU6476GetButton1(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte
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

'////////////  ve

CBW(15) = &HC4
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
 AU6476GetButton1 = 0
 Exit Function
End If

'2. Readdata stage
  
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

 For i = 0 To 39
  Debug.Print "i="; i; Hex(ReadData(i))
  
Next i
 

 

AU6476GetButton1 = ReadData(37)

 
 
  
      
    
   


'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
 
 
 
 If NumberOfBytesRead <> 13 Then
 
  result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 End If
 
 
 
If result = 0 Then
 AU6476GetButton1 = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
     AU6476GetButton1 = 0
 End If

 
End Function



Public Function Read_MS_Speed_AU6476E55(LBA As Long, Lun As Byte, CBWDataTransferLength As Long, BitWidth As String) As Byte
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
CBW(13) = &H0
CBW(14) = &H10
CBW(15) = &HC7
CBW(16) = &H1F
CBW(17) = &H5
CBW(18) = &H8F

'////////////// LUN
'CBW(13) = Lun                    '00

'///////////// CBD Len
'CBW(14) = &HA                '0a

'////////////  ve

CBW(19) = &HC4
CBW(20) = Lun * 32
LBAByte(0) = (LBA Mod 256)
tmpV(0) = Int(LBA / 256)
LBAByte(1) = (tmpV(0) Mod 256)
tmpV(1) = Int(tmpV(0) / 256)
LBAByte(2) = (tmpV(1) Mod 256)
tmpV(2) = Int((tmpV(1) / 256))
LBAByte(3) = (tmpV(2) Mod 256)

CBW(21) = LBAByte(3)         '00
CBW(22) = LBAByte(2)         '00
CBW(23) = LBAByte(1)         '00
CBW(24) = LBAByte(0)         '40

'/////////////  Reverve
CBW(25) = 0

'//////////// Transfer Len

TransferLen = Int(CBWDataTransferLength / 512)

TransferLenLSB = (TransferLen Mod 256)
tmpV(0) = Int(TransferLen / 256)
TransferLenMSB = (tmpV(0) / 256)

CBW(26) = TransferLenMSB      '00
CBW(27) = TransferLenLSB      '04

For i = 28 To 30
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
 Read_MS_Speed_AU6476E55 = 0
 Exit Function
End If

'2. Readdata stage
  
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

' For i = 0 To 15
' Debug.Print "i="; i; Hex(ReadData(i))
' Next

 

 

Read_MS_Speed_AU6476E55 = 0

If BitWidth = "8Bits" Then

  If (ReadData(16) = &HE0) Then ' &H70 : 8 bit, 50MHZ
      Read_MS_Speed_AU6476E55 = 1
      
  End If
  
End If
  
  
If BitWidth = "4Bits" Then
     Tester.Print "MS width="; Hex(ReadData(16))
  If (ReadData(16) = &H51) Or (ReadData(16) = &H40) Then  'it should be &H40 , but it is &HC1 in fact
      Read_MS_Speed_AU6476E55 = 1
       Tester.Print "MS is 4 bits, 40 MHZ"
   
      
  End If
  
End If
    
    
If Read_MS_Speed_AU6476E55 = 0 Then
Exit Function
End If
  


'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
 
 
 
 If NumberOfBytesRead <> 13 Then
 
  result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 End If
 
 
 
If result = 0 Then
 Read_MS_Speed_AU6476E55 = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
     Read_MS_Speed_AU6476E55 = 0
Else
     Read_MS_Speed_AU6476E55 = 1
   
End If

 
End Function

Public Function Read_MS_Speed_AU6476(LBA As Long, Lun As Byte, CBWDataTransferLength As Long, BitWidth As String) As Byte
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

'////////////  ve

CBW(15) = &HC4
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
 Read_MS_Speed_AU6476 = 0
 Exit Function
End If

'2. Readdata stage
  
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

'  For i = 0 To 15
 ' Debug.Print "i="; i; Hex(ReadData(i))
 ' Next

 

 

Read_MS_Speed_AU6476 = 0

If BitWidth = "1Bits" Then

  If (ReadData(16) = &H0) Then  ' &H70 : 8 bit, 50MHZ
      Read_MS_Speed_AU6476 = 1
      
  End If
  
End If
  


If BitWidth = "8Bits" Then

  If (ReadData(16) = &HE0) Then ' &H70 : 8 bit, 50MHZ
      Read_MS_Speed_AU6476 = 1
      
  End If
  
End If
  
  
If BitWidth = "4Bits" Then
     Tester.Print "MS width="; Hex(ReadData(16))
  If (ReadData(16) = &H51) Or (ReadData(16) = &H40) Then  'it should be &H40 , but it is &HC1 in fact
      Read_MS_Speed_AU6476 = 1
       Tester.Print "MS is 4 bits, 40 MHZ"
   
      
  End If
  
End If
    
    
If Read_MS_Speed_AU6476 = 0 Then
Exit Function
End If
  


'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
 
 
 
 If NumberOfBytesRead <> 13 Then
 
  result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 End If
 
 
 
If result = 0 Then
 Read_MS_Speed_AU6476 = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
     Read_MS_Speed_AU6476 = 0
Else
     Read_MS_Speed_AU6476 = 1
   
End If

 
End Function
 
Public Function Read_MS_Speed_AU6471(LBA As Long, Lun As Byte, CBWDataTransferLength As Long, BitWidth As String) As Byte
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

'////////////  ve

CBW(15) = &HC4
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
 Read_MS_Speed_AU6471 = 0
 Exit Function
End If

'2. Readdata stage
  
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

' For i = 0 To 15
' Debug.Print "i="; i; Hex(ReadData(i))
' Next

 

 

Read_MS_Speed_AU6471 = 0

If BitWidth = "8Bits" Then

  If (ReadData(16) = &HE0) Then ' &H70 : 8 bit, 50MHZ
      Read_MS_Speed_AU6471 = 1
      
  End If
  
End If
  
  
If BitWidth = "4Bits" Then
     Tester.Print "MS width="; Hex(ReadData(17))
  If (ReadData(17) = &H51) Then  'it should be &H40 , but it is &HC1 in fact
      Read_MS_Speed_AU6471 = 1
       Tester.Print "MS is 4 bits, 40 MHZ"
   
      
  End If
  
End If
    
    
If BitWidth = "1Bits" Then
     Tester.Print "MS width="; Hex(ReadData(17))
  If (ReadData(17) = &H0) Then  'it should be &H40 , but it is &HC1 in fact
      Read_MS_Speed_AU6471 = 1
       Tester.Print "MS is 1 bits, 20 MHZ"
   
      
  End If
  
End If
    
If Read_MS_Speed_AU6471 = 0 Then
Exit Function
End If
  


'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
 
 
 
 If NumberOfBytesRead <> 13 Then
 
  result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 End If
 
 
 
If result = 0 Then
 Read_MS_Speed_AU6471 = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
     Read_MS_Speed_AU6471 = 0
Else
     Read_MS_Speed_AU6471 = 1
   
End If

 
End Function
 
Public Function Read_MS_Speed(LBA As Long, Lun As Byte, CBWDataTransferLength As Long, BitWidth As String) As Byte
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

'////////////  ve

CBW(15) = &HC4
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
 Read_MS_Speed = 0
 Exit Function
End If

'2. Readdata stage
  
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

' For i = 0 To 15
' Debug.Print "i="; i; Hex(ReadData(i))
' Next

 

 

Read_MS_Speed = 0

If BitWidth = "8Bits" Then

  If (ReadData(16) = &HE0) Then ' &H70 : 8 bit, 50MHZ
      Read_MS_Speed = 1
      
  End If
  
End If
  
  
If BitWidth = "4Bits" Then
     Tester.Print "MS width="; Hex(ReadData(16))
  If (ReadData(16) = &HC1) Then  'it should be &HC0 , but it is &HC1 in fact
      Read_MS_Speed = 1
       Tester.Print "MS is 4 bits, 40 MHZ"
   
      
  End If
  
End If
    
    
If Read_MS_Speed = 0 Then
Exit Function
End If
  


'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
 
 
 
 If NumberOfBytesRead <> 13 Then
 
  result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 End If
 
 
 
If result = 0 Then
 Read_MS_Speed = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
     Read_MS_Speed = 0
Else
     Read_MS_Speed = 1
   
End If

 
End Function
Public Function Read_MS_Speed_AU6473(LBA As Long, Lun As Byte, CBWDataTransferLength As Long, BitWidth As String) As Byte
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

Dim HiByteValue As Byte

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

'////////////  ve

CBW(15) = &HC4
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
 Read_MS_Speed_AU6473 = 0
 Exit Function
End If

'2. Readdata stage
  
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

' For i = 0 To 15
' Debug.Print "i="; i; Hex(ReadData(i))
' Next

 

 

Read_MS_Speed_AU6473 = 0

If Lun = 0 Then
    
    HiByteValue = CAndValue(ReadData(17), &HF0)      'just compare HighByte
    
    If BitWidth = "8Bits" Then
        Tester.Print "MS width="; Hex(ReadData(17))
        If (HiByteValue = &HD0) Then
            
            Read_MS_Speed_AU6473 = 1
            Tester.Print "Lun0: MS is 8 bits, 40 MHZ"
        Else
            Tester.Print "Lun0: MS is NOT 8 bits, 40 MHZ"
        End If
  
    End If
  
    If BitWidth = "4Bits" Then
        Tester.Print "MS width="; Hex(ReadData(17))
        If (HiByteValue = &H50) Then
            Read_MS_Speed_AU6473 = 1
            Tester.Print "Lun0: MS is 4 bits, 40 MHZ"
        Else
            Tester.Print "Lun0: MS is NOT 4 bits, 40 MHZ"
        End If

    End If
    
    If BitWidth = "1Bit" Then
        Tester.Print "MS width="; Hex(ReadData(17))
        If (HiByteValue = &H1) Then
            Read_MS_Speed_AU6473 = 1
            Tester.Print "Lun0: MS is 1 bit, 40 MHZ"
        Else
            Tester.Print "Lun0: MS is NOT 1 bits, 40 MHZ"
        End If
    End If
    
Else    'LUN1

    HiByteValue = CAndValue(ReadData(23), &HF0)      'just compare HighByte
    If BitWidth = "8Bits" Then
        Tester.Print "MS width="; Hex(ReadData(23))
        If (HiByteValue = &HD0) Then
            Read_MS_Speed_AU6473 = 1
            Tester.Print "Lun1: MS is 8 bits, 40 MHZ"
        Else
            Tester.Print "Lun1: MS is NOT 8 bits, 40 MHZ"
        End If
  
    End If
  
    If BitWidth = "4Bits" Then
        Tester.Print "MS width="; Hex(ReadData(23))
        If (HiByteValue = &H50) Then
            Read_MS_Speed_AU6473 = 1
            Tester.Print "Lun1: MS is 4 bits, 40 MHZ"
        Else
            Tester.Print "Lun1: MS is NOT 4 bits, 40 MHZ"
        End If
    End If
    
    If BitWidth = "1Bit" Then
        Tester.Print "MS width="; Hex(ReadData(23))
        If (HiByteValue = &H1) Then
            Read_MS_Speed_AU6473 = 1
            Tester.Print "Lun1: MS is 1 bit, 40 MHZ"
        Else
            Tester.Print "Lun1: MS is NOT 1 bits, 40 MHZ"
        End If
    End If

End If
    
If Read_MS_Speed_AU6473 = 0 Then
    Exit Function
End If
  


'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
 
 
 
 If NumberOfBytesRead <> 13 Then
 
  result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 End If
 
 
 
If result = 0 Then
 Read_MS_Speed_AU6473 = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
     Read_MS_Speed_AU6473 = 0
Else
     Read_MS_Speed_AU6473 = 1
   
End If

 
End Function
Public Function RequestSense(Lun As Byte) As Byte

Dim CBW(0 To 30) As Byte
Dim ReadData(0 To 17) As Byte
Dim CSW(0 To 12) As Byte
Dim i As Integer
Dim NumberOfBytesWritten As Long
 
Dim NumberOfBytesRead As Long

    For i = 0 To 30
    
        CBW(i) = 0
    
    Next i
    
    
     For i = 0 To 17
    
        RequestSenseData(i) = 0
    
    Next i

CBW(0) = &H55 'signature
CBW(1) = &H53
CBW(2) = &H42
CBW(3) = &H43


CBW(4) = &H1  'package ID
CBW(5) = &H2
CBW(6) = &H3
CBW(7) = &H4


CBW(8) = &H12
CBW(9) = &H0
CBW(10) = &H0
CBW(11) = &H0





CBW(12) = &H80 '    CBW FLAG 0000
CBW(13) = Lun
 
CBW(14) = &HA
CBW(15) = &H3

CBW(16) = Lun * 32
CBW(19) = &H12

'1. CBW
RequestSense = 0

result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)    'out
 
 
If result = 0 Then
    RequestSense = 0
    Exit Function
End If
 

'2. Request Data input
result = ReadFile _
         (ReadHandle, _
          RequestSenseData(0), _
          18, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
          
          
          
If result = 0 Then
    RequestSense = 0
    Exit Function
End If
           
          
'3. CSW
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
        
If result = 0 Then
    RequestSense = 0
    Exit Function
End If
 

If CSW(12) = 1 Then
    RequestSense = 0
Else
    RequestSense = 1
End If

End Function

Public Function TestUnitSpeed(Lun As Byte) As Byte
Dim CBW(0 To 30) As Byte
Dim CSW(0 To 12) As Byte
Dim i As Integer
Dim NumberOfBytesWritten As Long
Dim NumberOfBytesRead As Long
Dim result As Long

     For i = 0 To 30
    
        CBW(i) = 0
    
    Next i

CBW(0) = &H55 'signature
CBW(1) = &H53
CBW(2) = &H42
'CBW(3) = &H43  ' orgioinal tag
CBW(3) = &H55  ' for test unit speed


CBW(4) = &H1  'package ID
CBW(5) = &H2
CBW(6) = &H3
CBW(7) = &H4


CBW(8) = &H0
CBW(9) = &H0
CBW(10) = &H0
CBW(11) = &H0


CBW(12) = &H80 '    CBW FLAG 0000
'CBW(13) = &H0
CBW(13) = Lun
CBW(14) = &H0
CBW(15) = &H0
CBW(16) = Lun * 32


'1. CBW output

TestUnitSpeed = 0

result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)

If result = 0 Then
     TestUnitSpeed = 0
    Exit Function
End If


'2. CSW input
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
'Debug.Print "TestUnit Speed"
'For i = 0 To 12
'Debug.Print i, CSW(i)
'Next i

If result = 0 Then
    TestUnitSpeed = 0
   Exit Function
End If

 
'3 CSW Status
If CSW(12) = 1 Then
    TestUnitSpeed = 0
    
    Else
    TestUnitSpeed = 1
    End If

End Function

 
 Public Sub AU6254Test(rv0 As Byte, rv1 As Byte, rv2 As Byte, rv3 As Byte, rv4 As Byte)
 On Error Resume Next
 Dim TestCounter As Integer
 Dim TimeInterval
 Dim HubEmuCounter As Integer
 Dim ContinueRun As Integer
 Dim PwrTime As Single
 Dim Au6254Speed As Integer
 
 ContinueRun = 1
 PwrTime = 0.8
 
 
 Tester.txtmsg.Text = ""
              
                If PCI7248InitFinish = 0 Then
                  PCI7248ExistAU6254
                End If
                          
               '====================================
               '            Hub exist test3
               ' ====================================
               '****************************
                'For HubEmuCounter = 1 To 5
               '****************************
                    WinExec "off.exe", 0
                    CardResult = DO_WritePort(card, Channel_P1CL, &H4)
                '  CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                 ' Print "1"
                  Call MsecDelay(0.3) ' system let hub driver unload time
               
                  '============= dETECT FOR Disconnect fail
                   
                '  rv0 = AU6254_GetDevice(0, 1, "6254")
                '  If rv0 = 1 Then
                '     MsgBox " Chip Disconnect fail, wait for shutdown"
                     
                '      Do
                '      DoEvents
                '      Loop While (1)
                '  End If
                  
                  Tester.Cls
                  Tester.Print "2"
                    WinExec "on.exe", 0
                  CardResult = DO_WritePort(card, Channel_P1A, &H1E)
                  Call MsecDelay(3) ' Hub Power on
                  If CardResult <> 0 Then
                      MsgBox "Power on fail"
                      End
                  End If
                  ReaderExist = 0
                  rv0 = 0
                  OldTimer = Timer
                  Tester.Cls
                  Tester.Print "3"
                  Do
                      Call MsecDelay(0.2)
                      DoEvents
                       rv0 = AU6254_GetDevice(0, 1, "6254")
                       
                      TimeInterval = Timer - OldTimer
                  Loop While rv0 = 0 And TimeInterval < 5
                 
                  Tester.Print "rv0 ="; rv0; "TimeInterval="; TimeInterval
                  
                 '****************************
              '  Next HubEmuCounter
              '  MsgBox "Ok"
              '  Exit Sub
                '****************************
                
                If rv0 <> 1 Then 'Hub unknow
                     rv1 = 4
                     rv2 = 4
                     rv3 = 4
                     rv4 = 4
                     If ContinueRun = 0 Then
                        Exit Sub
                     End If
                End If
        
                Tester.Print "rv0="; rv0; "--- Hub Exist"
    
                '==========================  Hub Detect
                '1. usb 2.0 reader
                '2. usb 1.0 flash
                '3, usb 6610
                '4. usb keyboard
        
               '======== PWR ctrl
                CardResult = DO_WritePort(card, Channel_P1A, &H1C)
                Call MsecDelay(PwrTime)
                
                '===============   must at here to test speed , otherwise driver will overlap the test result
                
                HubPort = 0
                ReaderExist = 0
                rv1 = 0
                ClosePipe
                'rv1 = AU6610Test
                rv1 = CBWTest_New_AU9254(0, 1, "6335")
                ClosePipe
                Au6254Speed = UsbSpeedTestResult
                 
                Tester.Print "rv1="; rv1; "--- 2.0 speed and isochrous pipe test"
                
                
                   
               If rv1 <> 1 Then
                   
                    rv2 = 4
                    rv3 = 4
                    rv4 = 4
                    
                    If rv1 >= 2 Then  ' speed error
                         Exit Sub
                    End If
                    
                    If ContinueRun = 0 Then
                        Exit Sub
                    End If
                End If
                
                 Tester.Print "rv1="; rv1; "--- 2.0 speed and isochrous pipe test"
               
               
                CardResult = DO_WritePort(card, Channel_P1A, &H18)
                Call MsecDelay(PwrTime)
                CardResult = DO_WritePort(card, Channel_P1A, &H10)
                Call MsecDelay(PwrTime)
                CardResult = DO_WritePort(card, Channel_P1A, &H0)
                Call MsecDelay(PwrTime)
                
                '===== 6335 test usb 2.0 speed
              
                   
                 
                
                '****************************
              '  Next HubEmuCounter
              '  MsgBox "Ok"
              '  Exit Sub
                '****************************
             
                 
                 
               '=================================
               '   test usb 1.1 flash
               '================================
                HubPort = 3
                LBA = LBA + 1
                ReaderExist = 0
                ClosePipe
                  rv2 = CBWTest_New_AU9254(0, 1, "1758")
                ClosePipe
           
                Tester.Print "rv2="; rv2; "--- 2.0 speed and Bulk r/w"
                Tester.Print " UsbSpeedTestResult="; UsbSpeedTestResult; "---usb speed error r/w"
                
                
                If rv2 <> 1 Then
                
                   
             
                   rv3 = 4
                   rv4 = 4
                   If ContinueRun = 0 Then
                        Exit Sub
                    End If
                End If
                 
                
               '=================================
               '   test 6610 for isochrous pipe
               '================================
               
  
               
                HubPort = 1
                ReaderExist = 0
                ClosePipe
                rv3 = CBWTest_New_AU9254(0, 1, "9369")
                ClosePipe
              
             
             
                Tester.Print "rv3="; rv3; "--- 1.1 speed and isochrorous r/w"
               
               
                 CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                 If rv3 <> 1 Then
                        
                    
                       
                      rv4 = 4
                    If ContinueRun = 0 Then
                        Exit Sub
                    End If
                     
                  Else
                  
                     ' If LightON <> 243 Then
                     '     rv3 = 2
                      '   End If
                  
                     
                  End If
                
                
                 
                
                
                
               '=================================
               '   test usb 1.1 key board
               '================================
                  Call MsecDelay(1)
                 Tester.txtmsg.Text = ""
                 Tester.txtmsg.SetFocus
                 
                 HubPort = 2
                 ReaderExist = 0
                 ClosePipe
                   rv4 = CBWTest_New_AU9254(0, 1, "9462")
                 ClosePipe
                
                ' ReaderExist = 0
                ' ClosePipe
                '   rv4 = CBWTest_New_AU9254(0, 1, "9462")
                ' ClosePipe
        
                  '  CardResult = DO_WritePort(card, Channel_P1A, &HE)    ' usb 2.0 falsh
                
              '  Call MsecDelay(4)
                
                ' keybaord control
                 '  CardResult = DO_ReadPort(card, Channel_P1B, LightON)
                 '  Call MsecDelay(0.1)
                '   CardResult = DO_WritePort(card, Channel_P1CH, &H0)
                 '   Call MsecDelay(0.3)
                 
                 ' CardResult = DO_WritePort(card, Channel_P1CH, &H1)
                  '    Call MsecDelay(1.2)
                
                 
                    
                 ' CardResult = DO_WritePort(card, Channel_P1CH, &H0)
                 '   Call MsecDelay(0.2)
                  'CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' close power
                   
                 ' Tester.print " LightON="; LightON
                  
                  
                   
                   
                   
               '    If InStr(txtmsg.Text, "..") <> 0 Then
                  
                '      rv4 = 1
                    
                '      If LightON = 238 Then
                '      rv4 = 1
                     
                 '    Else
                 '     Tester.print "GPO Fail"
                     
                  '   End If
                     
               '  Else
               '    Tester.print "Keyboard fail"
                    
                 
               '  End If
               
                  
                  'If rv4 <> 1 Then
                  '    rv4 = 2
                  ' End If
                   Tester.Print "rv4="; rv4; "--- 12 MHZ speed and interrupt  r/w"
                
                
                    
                 If rv0 = 1 Or rv1 = 1 Or rv2 = 1 Or rv3 = 1 Or rv4 = 1 Then
                 
                 
                    If rv0 * rv1 * rv2 * rv3 * rv4 = 1 Then
                      Rv1ContinueFail = 0
                    End If
                    
                  
                    If rv0 <> 1 Then
                    'Rv1ContinueFail = 0
                     ReaderExist = 0
                     rv0 = AU6254_GetDevice(0, 1, "6254")
                    End If
                    
                  If rv1 <> 1 Then
                  
                  
                   
                  
                  
                
                    HubPort = 0
                    ReaderExist = 0
                    ClosePipe
                    'rv1 = AU6610Test
                    rv1 = CBWTest_New_AU9254(0, 1, "6335")
                    ClosePipe
                     Au6254Speed = UsbSpeedTestResult
                     
                    
                   
                   
                     
                     
                  End If
                  
                  
                  If rv2 <> 1 Then
                    ' Rv1ContinueFail = 0
                     HubPort = 3
                     ReaderExist = 0
                    ClosePipe
                    rv2 = CBWTest_New_AU9254(0, 1, "1758")
                    ClosePipe
                  End If
                  
                  If rv3 <> 1 Then
                    'Rv1ContinueFail = 0
                    HubPort = 1
                    ReaderExist = 0
                      ClosePipe
                     rv3 = CBWTest_New_AU9254(0, 1, "9369")
                    ClosePipe
              
                  End If
                  
                  If rv4 <> 1 Then
                   '  Rv1ContinueFail = 0
                     HubPort = 2
                     ReaderExist = 0
                     ClosePipe
                     rv4 = CBWTest_New_AU9254(0, 1, "9462")
                     ClosePipe
        
                  End If
                  
                  
                    If rv1 + rv2 + rv3 + rv4 = 0 Then
                     
                      Rv1ContinueFail = Rv1ContinueFail + 1
                     End If
                      
                     
                     
                    If Rv1ContinueFail > 1 Then  ' use another AU6335 plug into to solve the hub hang probelm
                    CardResult = DO_WritePort(card, Channel_P1CL, &H0)
                    Call MsecDelay(2.5)
                    CardResult = DO_WritePort(card, Channel_P1CL, &H4)
                    Call MsecDelay(2)
                   End If
                 
                Tester.Print "RTrv0="; rv0
                Tester.Print "RTrv1="; rv1
                Tester.Print "RTrv2="; rv2
                Tester.Print "RTrv3="; rv3
                Tester.Print "RTrv4="; rv4

                
             End If
             
             
            '========== binning
             If rv0 = 0 Then   ' hub unknow  ,bin2
               rv1 = 4
               rv2 = 4
               rv3 = 4
               rv4 = 4
               AU6254TestMsg = "Hub Unknow"
               Exit Sub
          End If
          
        
               If rv0 = 1 Then   ' hub unknow  , bin3
                 If Au6254Speed = 2 Then     ' hub speed
                   rv1 = 2
                   rv2 = 4
                   rv3 = 4
                   rv4 = 4
                    AU6254TestMsg = "USB 2.0 speed error"
                   Exit Sub
                End If
                
                If rv1 * rv2 * rv3 * rv4 = 0 Then  ' bin4 down stream port unknow device
                     PortFail = ""
                     If rv1 = 0 Then
                       PortFail = PortFail & "port1 unknow,"
                    End If
                    
                       If rv2 = 0 Then
                       PortFail = PortFail & "port2 unknow,"
                    End If
                    
                       If rv3 = 0 Then
                       PortFail = PortFail & "port3 unknow,"
                    End If
                    
                       If rv4 = 0 Then
                       PortFail = PortFail & "port4 unknow,"
                    End If
                      
                      
                    
                    rv1 = 1
                    rv2 = 1
                    rv3 = 2
                    rv4 = 4
                    
                    AU6254TestMsg = PortFail
                    
                    Exit Sub
                End If
                
                
                
                
                If rv2 >= 2 Then   ' bin5 down stream port unknow device
                    rv1 = 1
                    rv2 = 1
                    rv3 = 1
                    rv4 = 2
                    AU6254TestMsg = "2.0 Reader SD R/W fail"
                    Exit Sub
                End If
                
                  
                If rv3 >= 2 Then   ' bin5 down stream port unknow device
                    rv1 = 1
                    rv2 = 1
                    rv3 = 1
                    rv4 = 2
                    AU6254TestMsg = "1.0 Reader SD R/W fail"
                    Exit Sub
                End If
                
                
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                
                 
               If LightOn <> 225 And LightOn <> 224 Then ' bin5 : light on fail
                   rv1 = 1
                   rv2 = 1
                   rv3 = 1
                   rv4 = 2
                   AU6254TestMsg = "GPO R/W fail"
               End If
                Exit Sub
            End If
            
                


 End Sub

Public Function Write_Data_AU6377(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte

Dim CBW(0 To 30) As Byte
Dim CSW(0 To 12) As Byte
Dim NumberOfBytesWritten As Long
Dim NumberOfBytesRead As Long
Dim CBWDataTransferLen(0 To 3) As Byte
Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim i As Integer
Dim tmpV(0 To 2) As Long
Dim opcode As Byte

opcode = &H2A
'Buffer(0) = &H33 'CByte(Text2.Text)
'Buffer(1) = &H44


    For i = 0 To 30
    
        CBW(i) = 0
    
    Next i
    
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
CBW(12) = &H0                 '80

'////////////// LUN
CBW(13) = Lun                    '00

'///////////// CBD Len
CBW(14) = &HA                '0a

'////////////  UFI command

CBW(15) = opcode
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

'Tester.print Hex(CBW(17)); " "; Hex(CBW(18)); " "; Hex(CBW(19)); " "; Hex(CBW(20))
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

 
'1. CBW output
 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)    'out

If result = 0 Then
    Write_Data_AU6377 = 0
    Exit Function
End If
 
 
 
'2, Output data
result = WriteFile _
       (WriteHandle, _
       Pattern_AU6377(0), _
       CBWDataTransferLength, _
       NumberOfBytesWritten, _
       0)    'out

 
If result = 0 Then
    Write_Data_AU6377 = 0
    Exit Function
End If

'3 . CSW
result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
        
If result = 0 Then
    Write_Data_AU6377 = 0
    Exit Function
End If
 
 
 
If CSW(12) = 1 Then
Write_Data_AU6377 = 0

Else
Write_Data_AU6377 = 1
End If
End Function

Public Function Write_Data_AU6371(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte

Dim CBW(0 To 30) As Byte
Dim CSW(0 To 12) As Byte
Dim NumberOfBytesWritten As Long
Dim NumberOfBytesRead As Long
Dim CBWDataTransferLen(0 To 3) As Byte
Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim i As Integer
Dim tmpV(0 To 2) As Long
Dim opcode As Byte

opcode = &H2A
'Buffer(0) = &H33 'CByte(Text2.Text)
'Buffer(1) = &H44


    For i = 0 To 30
    
        CBW(i) = 0
    
    Next i
    
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
CBW(12) = &H0                 '80

'////////////// LUN
CBW(13) = Lun                    '00

'///////////// CBD Len
CBW(14) = &HA                '0a

'////////////  UFI command

CBW(15) = opcode
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

'Tester.print Hex(CBW(17)); " "; Hex(CBW(18)); " "; Hex(CBW(19)); " "; Hex(CBW(20))
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

 
'1. CBW output
 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)    'out

If result = 0 Then
    Write_Data_AU6371 = 0
    Exit Function
End If
 
 
 
'2, Output data
result = WriteFile _
       (WriteHandle, _
       AU6371Pattern(0), _
       CBWDataTransferLength, _
       NumberOfBytesWritten, _
       0)    'out

 
If result = 0 Then
    Write_Data_AU6371 = 0
    Exit Function
End If

'3 . CSW
result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
        
If result = 0 Then
    Write_Data_AU6371 = 0
    Exit Function
End If
 
 
 
If CSW(12) = 1 Then
Write_Data_AU6371 = 0

Else
Write_Data_AU6371 = 1
End If
End Function
Public Function Write_Data_AU6375(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte

Dim CBW(0 To 30) As Byte
Dim CSW(0 To 12) As Byte
Dim NumberOfBytesWritten As Long
Dim NumberOfBytesRead As Long
Dim CBWDataTransferLen(0 To 3) As Byte
Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim i As Integer
Dim tmpV(0 To 2) As Long
Dim opcode As Byte

opcode = &H2A
'Buffer(0) = &H33 'CByte(Text2.Text)
'Buffer(1) = &H44


    For i = 0 To 30
    
        CBW(i) = 0
    
    Next i
    
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
CBW(12) = &H0                 '80

'////////////// LUN
CBW(13) = Lun                    '00

'///////////// CBD Len
CBW(14) = &HA                '0a

'////////////  UFI command

CBW(15) = opcode
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

'Tester.print Hex(CBW(17)); " "; Hex(CBW(18)); " "; Hex(CBW(19)); " "; Hex(CBW(20))
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

 
'1. CBW output
 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)    'out

If result = 0 Then
    Write_Data_AU6375 = 0
    Exit Function
End If
 
 
 
'2, Output data
result = WriteFile _
       (WriteHandle, _
       Pattern_AU6375(0), _
       CBWDataTransferLength, _
       NumberOfBytesWritten, _
       0)    'out

 
If result = 0 Then
    Write_Data_AU6375 = 0
    Exit Function
End If

'3 . CSW
result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
        
If result = 0 Then
    Write_Data_AU6375 = 0
    Exit Function
End If
 
 
 
If CSW(12) = 1 Then
Write_Data_AU6375 = 0

Else
Write_Data_AU6375 = 1
End If
End Function

Public Function AU6254_GetDevice(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

   CBWDataTransferLength = 1024
 
'   For i = 0 To CBWDataTransferLength - 1
    
'         ReadData(i) = 0

'   Next

    If PreSlotStatus <> 1 Then
        AU6254_GetDevice = 4
        Exit Function
    End If
    '========================================
   
    AU6254_GetDevice = 0
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
            TmpString = GetDeviceNameHub(Vid_PID)
        Loop While TmpString = "" And TimerCounter < 10
    End If
    '=======================================
    If ReaderExist = 0 And TmpString <> "" Then
      ReaderExist = 1
    End If
    '=======================================
    If ReaderExist = 0 And TmpString = "" Then
      AU6254_GetDevice = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
 
     
    AU6254_GetDevice = 1
        
    
    End Function

Public Function SetOverCurrent(PreviousResult As Byte) As Byte

If PreviousResult <> 1 Then
SetOverCurrent = 4
Exit Function
End If

Dim CBW(0 To 30) As Byte
Dim CSW(0 To 12) As Byte
Dim i As Integer
Dim NumberOfBytesWritten As Long
Dim NumberOfBytesRead As Long
Dim result As Long

     For i = 0 To 30
    
        CBW(i) = 0
    
    Next i

CBW(0) = &H55 'signature
CBW(1) = &H53
CBW(2) = &H42
CBW(3) = &H43


CBW(4) = &H1  'package ID
CBW(5) = &H2
CBW(6) = &H3
CBW(7) = &H4


CBW(8) = &H0
CBW(9) = &H0
CBW(10) = &H0
CBW(11) = &H0


CBW(12) = &H80 '    CBW FLAG 0000
'CBW(13) = &H0
CBW(13) = 0
CBW(14) = &H0
CBW(15) = &HC7
CBW(16) = &H1
CBW(17) = &H1
 

'1. CBW output

SetOverCurrent = 0

result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)

'If result = 0 Then
'    TestUnitReady = 0
'    Exit Function
'End If


'2. CSW input
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

If result = 0 Then
    SetOverCurrent = 0
    Exit Function
End If

 
'3 CSW Status
If CSW(12) = 1 Then
    SetOverCurrent = 0
    
    Else
    SetOverCurrent = 1
End If

End Function

Public Function SetOverCurrent5(PreviousResult As Byte) As Byte

If PreviousResult <> 1 Then
SetOverCurrent5 = 4
Exit Function
End If

Dim CBW(0 To 30) As Byte
Dim CSW(0 To 12) As Byte
Dim i As Integer
Dim NumberOfBytesWritten As Long
Dim NumberOfBytesRead As Long
Dim result As Long

     For i = 0 To 30
    
        CBW(i) = 0
    
    Next i

CBW(0) = &H55 'signature
CBW(1) = &H53
CBW(2) = &H42
CBW(3) = &H43


CBW(4) = &H1  'package ID
CBW(5) = &H2
CBW(6) = &H3
CBW(7) = &H4


CBW(8) = &H0
CBW(9) = &H0
CBW(10) = &H0
CBW(11) = &H0


CBW(12) = &H80 '    CBW FLAG 0000
'CBW(13) = &H0
CBW(13) = 0
CBW(14) = &H0
CBW(15) = &HC7
CBW(16) = &H1
CBW(17) = &H1
 

'1. CBW output

result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in


SetOverCurrent5 = 0

result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)

'If result = 0 Then
'    TestUnitReady = 0
'    Exit Function
'End If


'2. CSW input
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

If result = 0 Then
    SetOverCurrent5 = 0
    Exit Function
End If

 
'3 CSW Status
If CSW(12) = 1 Then
    SetOverCurrent5 = 0
    
    Else
    SetOverCurrent5 = 1
End If

End Function


Public Function ReInitial(Lun As Byte) As Byte
Dim CBW(0 To 30) As Byte
Dim CSW(0 To 12) As Byte
Dim i As Integer
Dim NumberOfBytesWritten As Long
Dim NumberOfBytesRead As Long
Dim result As Long

     For i = 0 To 30
    
        CBW(i) = 0
    
    Next i

CBW(0) = &H55 'signature
CBW(1) = &H53
CBW(2) = &H42
CBW(3) = &H43


CBW(4) = &H1  'package ID
CBW(5) = &H2
CBW(6) = &H3
CBW(7) = &H4


CBW(8) = &H0
CBW(9) = &H0
CBW(10) = &H0
CBW(11) = &H0


CBW(12) = &H80 '    CBW FLAG 0000
'CBW(13) = &H0
CBW(13) = Lun
CBW(14) = &H0
CBW(15) = &HC7
CBW(16) = &H3
CBW(17) = &H30
CBW(18) = &H35
CBW(19) = &H38
CBW(20) = &H46
CBW(21) = &H9F


'1. CBW output

ReInitial = 0

result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)

'If result = 0 Then
'    TestUnitReady = 0
'    Exit Function
'End If


'2. CSW input
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

If result = 0 Then
    ReInitial = 0
    Exit Function
End If

 
'3 CSW Status
If CSW(12) = 1 Then
    ReInitial = 0
    
    Else
    ReInitial = 1
End If

End Function

Public Function TestUnitReadyReadOverCurrent3(Lun As Byte) As Byte
Dim CBW(0 To 30) As Byte
Dim CSW(0 To 12) As Byte
Dim i As Integer
Dim NumberOfBytesWritten As Long
Dim NumberOfBytesRead As Long
Dim result As Long

     For i = 0 To 30
    
        CBW(i) = 0
    
    Next i

CBW(0) = &H55 'signature
CBW(1) = &H53
CBW(2) = &H42
CBW(3) = &H43


CBW(4) = &H1  'package ID
CBW(5) = &H2
CBW(6) = &H3
CBW(7) = &H4


CBW(8) = &H0
CBW(9) = &H0
CBW(10) = &H0
CBW(11) = &H0


CBW(12) = &H80 '    CBW FLAG 0000
'CBW(13) = &H0
CBW(13) = Lun
CBW(14) = &H0
CBW(15) = &H0
CBW(16) = Lun * 32

result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
 For i = 1 To 1
  Call MsecDelay(0.1)
 CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
 
 Tester.Print "1overCurrent value="; Hex(LightOff); " : 80:pass;00< 3.17:C0>3.43"
        If LightOff <> 128 Then
      '      Exit For
        End If
Next
'1. CBW output

TestUnitReadyReadOverCurrent3 = 7
 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)
 
 If result = 0 Then
     TestUnitReadyReadOverCurrent3 = 7
    Exit Function
 End If
'result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
Call MsecDelay(0.1)
 For i = 1 To 8
  Call MsecDelay(0.2)
 CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
' Tester.Print "2overCurrent value="; Hex(LightOFF); ""
 LightOff = CAndValue(LightOff, &HC0)
 Tester.Print "overCurrent value="; Hex(LightOff); " : 80:pass;00< 3.17:C0>3.43"
        If LightOff <> 128 Then
      '      Exit For
        End If
Next
  
Exit Function
 
  If LightOff = 128 Then  ' clamp between 3.45 and 3.15
  
      Call PowerSet2(1, "2.5", "0.5", 1, "2.5", "0.5", 1)
       Call MsecDelay(0.8)
      CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
      Tester.Print "check low resistor ="; Hex(LightOff); " C0 :pass"
      If LightOff = &HC0 Then
          
          Call PowerSet2(1, "3.5", "0.5", 1, "3.5", "0.5", 1)
          Call MsecDelay(0.8)
          CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
           Tester.Print "check high resistor ="; Hex(LightOff); " 00 :pass"
           If LightOff = 0 Then
           TestUnitReadyReadOverCurrent3 = 5
            End If
       End If
  
  
     
    Exit Function
 End If
 
  If LightOff = 0 Then
     TestUnitReadyReadOverCurrent3 = 6
    Exit Function
 End If
 

 

 
 
 
'Call MsecDelay(0.2)
'2. CSW input
 'result = ReadFile _
 '        (ReadHandle, _
 '         CSW(0), _
 '         13, _
 '         NumberOfBytesRead, _
 '         HIDOverlapped)  'in
 
 
Exit Function

If result = 0 Then
Tester.Print "5"
  '  TestUnitReadyReadOverCurrent3 = 0
  '  result = ReadFile _
  '       (ReadHandle, _
  '        CSW(0), _
  '        13, _
  '        NumberOfBytesRead, _
  '        HIDOverlapped)  'in
          
  '   result = ReadFile _
  '       (ReadHandle, _
  '        CSW(0), _
  '        13, _
  '        NumberOfBytesRead, _
  '        HIDOverlapped)  'in
          
    Exit Function
End If

 
'3 CSW Status
If CSW(12) = 1 Then
    TestUnitReadyReadOverCurrent3 = 0
    
    Else
    TestUnitReadyReadOverCurrent3 = 1
End If

End Function
Public Function TestUnitReadyReadOverCurrentA2D4(Lun As Byte) As Byte
Dim CBW(0 To 30) As Byte
Dim CSW(0 To 12) As Byte
Dim i As Integer
Dim NumberOfBytesWritten As Long
Dim NumberOfBytesRead As Long
Dim result As Long

     For i = 0 To 30
    
        CBW(i) = 0
    
    Next i

CBW(0) = &H55 'signature
CBW(1) = &H53
CBW(2) = &H42
CBW(3) = &H43


CBW(4) = &H1  'package ID
CBW(5) = &H2
CBW(6) = &H3
CBW(7) = &H4


CBW(8) = &H0
CBW(9) = &H0
CBW(10) = &H0
CBW(11) = &H0


CBW(12) = &H80 '    CBW FLAG 0000
'CBW(13) = &H0
CBW(13) = Lun
CBW(14) = &H0
CBW(15) = &H0
CBW(16) = Lun * 32


'1. CBW output

TestUnitReadyReadOverCurrentA2D4 = 7
 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)
 
 If result = 0 Then
     TestUnitReadyReadOverCurrentA2D4 = 7
    Exit Function
 End If
 Call MsecDelay(1#)
 '====================================
 '   A2D begin
 '====================================
   
 Call DACGetData1Channel

 
  dma2205.Show
 dma2205.ShowDataAU6476 0, 100
 
 Dim tmp As Single
 For i = 0 To 99
 tmp = CSng(ADCBuffer(i) + 1) / CSng(32768) * 5#

 If tmp > AU6476Upper Then
       TestUnitReadyReadOverCurrentA2D4 = 6
  Exit Function
 End If
 
 If tmp < AU6476Lower Then
      TestUnitReadyReadOverCurrentA2D4 = 5
 Exit Function
 End If
 
 Next i
 
 
 

  
'3 CSW Status
    
    
'      result = ReadFile _
'         (ReadHandle, _
'          CSW(0), _
'          13, _
'          NumberOfBytesRead, _
'          HIDOverlapped)  'in

'If NumberOfBytesRead <> 13 Then
    
'    result = ReadFile _
'         (ReadHandle, _
'          CSW(0), _
'          13, _
'          NumberOfBytesRead, _
'          HIDOverlapped)  'in
' End If

 
'3 CSW Status
 
  TestUnitReadyReadOverCurrentA2D4 = 1
 

End Function
Public Function TestUnitReadyReadOverCurrentA2D(Lun As Byte) As Byte
Dim CBW(0 To 30) As Byte
Dim CSW(0 To 12) As Byte
Dim i As Integer
Dim NumberOfBytesWritten As Long
Dim NumberOfBytesRead As Long
Dim result As Long

     For i = 0 To 30
    
        CBW(i) = 0
    
    Next i

CBW(0) = &H55 'signature
CBW(1) = &H53
CBW(2) = &H42
CBW(3) = &H43


CBW(4) = &H1  'package ID
CBW(5) = &H2
CBW(6) = &H3
CBW(7) = &H4


CBW(8) = &H0
CBW(9) = &H0
CBW(10) = &H0
CBW(11) = &H0


CBW(12) = &H80 '    CBW FLAG 0000
'CBW(13) = &H0
CBW(13) = Lun
CBW(14) = &H0
CBW(15) = &H0
CBW(16) = Lun * 32


'1. CBW output

TestUnitReadyReadOverCurrentA2D = 7
 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)
 
 If result = 0 Then
     TestUnitReadyReadOverCurrentA2D = 7
    Exit Function
 End If
 Call MsecDelay(1#)
 '====================================
 '   A2D begin
 '====================================
   
 Call DACGetData1Channel

 
  dma2205.Show
 dma2205.ShowDataAU6476 0, 100
 
 Dim tmp As Single
 For i = 0 To 99
         tmp = CSng(ADCBuffer(i) + 1) / CSng(32768) * 5#

 
 If tmp > AU6476Upper Then
       TestUnitReadyReadOverCurrentA2D = 6
  Exit Function
 End If
 
 If tmp < AU6476Lower Then
      TestUnitReadyReadOverCurrentA2D = 5
 Exit Function
 End If
 
 Next i
 
 
 

   
 
'3 CSW Status
 
  TestUnitReadyReadOverCurrentA2D = 1
 

End Function
Public Function TestUnitReady5(Lun As Byte) As Byte
Dim CBW(0 To 30) As Byte
Dim CSW(0 To 12) As Byte
Dim i As Integer
Dim NumberOfBytesWritten As Long
Dim NumberOfBytesRead As Long
Dim result As Long

     For i = 0 To 30
    
        CBW(i) = 0
    
    Next i

CBW(0) = &H55 'signature
CBW(1) = &H53
CBW(2) = &H42
CBW(3) = &H43


CBW(4) = &H1  'package ID
CBW(5) = &H2
CBW(6) = &H3
CBW(7) = &H4


CBW(8) = &H0
CBW(9) = &H0
CBW(10) = &H0
CBW(11) = &H0


CBW(12) = &H80 '    CBW FLAG 0000
'CBW(13) = &H0
CBW(13) = Lun
CBW(14) = &H0
CBW(15) = &H0
CBW(16) = Lun * 32

result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
'1. CBW output

TestUnitReady5 = 0

result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)

'If result = 0 Then
'    TestUnitReady5 = 0
'    Exit Function
'End If


'2. CSW input
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

If result = 0 Then
    TestUnitReady5 = 0
    
     result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

    
   ' Exit Function
End If

 
'3 CSW Status
If CSW(12) = 1 Then
    TestUnitReady5 = 0
    
    Else
    TestUnitReady5 = 1
End If

End Function

Public Function TestUnitReady(Lun As Byte) As Byte
Dim CBW(0 To 30) As Byte
Dim CSW(0 To 12) As Byte
Dim i As Integer
Dim NumberOfBytesWritten As Long
Dim NumberOfBytesRead As Long
Dim result As Long

     For i = 0 To 30
    
        CBW(i) = 0
    
    Next i

CBW(0) = &H55 'signature
CBW(1) = &H53
CBW(2) = &H42
CBW(3) = &H43


CBW(4) = &H1  'package ID
CBW(5) = &H2
CBW(6) = &H3
CBW(7) = &H4


CBW(8) = &H0
CBW(9) = &H0
CBW(10) = &H0
CBW(11) = &H0


CBW(12) = &H80 '    CBW FLAG 0000
'CBW(13) = &H0
CBW(13) = Lun
CBW(14) = &H0
CBW(15) = &H0
CBW(16) = Lun * 32


'1. CBW output

TestUnitReady = 0

result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)

'If result = 0 Then
'    TestUnitReady = 0
'    Exit Function
'End If


'2. CSW input
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

If result = 0 Then
    TestUnitReady = 0
    
    Exit Function
End If

 
'3 CSW Status
If CSW(12) = 1 Then
    TestUnitReady = 0
    
    Else
    TestUnitReady = 1
End If

End Function

Public Function AU6476GetButton(Lun As Byte) As Byte
Dim CBW(0 To 30) As Byte
Dim CSW(0 To 12) As Byte
Dim i As Integer
Dim NumberOfBytesWritten As Long
Dim NumberOfBytesRead As Long
Dim result As Long

     For i = 0 To 30
    
        CBW(i) = 0
    
    Next i

CBW(0) = &H55 'signature
CBW(1) = &H53
CBW(2) = &H42
CBW(3) = &H43


CBW(4) = &H1  'package ID
CBW(5) = &H2
CBW(6) = &H3
CBW(7) = &H4


CBW(8) = &H0
CBW(9) = &H0
CBW(10) = &H0
CBW(11) = &H0


CBW(12) = &H80 '    CBW FLAG 0000
'CBW(13) = &H0
CBW(13) = Lun
CBW(14) = &H0
CBW(15) = &HC7
CBW(16) = &H13
CBW(17) = &H0
CBW(18) = &H8
CBW(19) = &HF

'1. CBW output

AU6476GetButton = 0

result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)

'If result = 0 Then
'    AU6476GetButton = 0
'    Exit Function
'End If


'2. CSW input
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

If result = 0 Then
    AU6476GetButton = 0
    Exit Function
End If

 
'3 CSW Status
If CSW(12) = 1 Then
    AU6476GetButton = 0
    
    Else
    AU6476GetButton = 1
End If

End Function
Public Function OpenPipe() As Byte
Dim WritePathName As String
Dim ReadPathName As String

OpenPipe = 0
WritePathName = Left(DevicePathName, Len(DevicePathName) - 2) & "\PIPE0"   '
'Debug.Print Lba
'Debug.Print "WritePathName="; WritePathName
WriteHandle = CreateFile _
             (WritePathName, _
            GENERIC_READ Or GENERIC_WRITE, _
            (FILE_SHARE_READ Or FILE_SHARE_WRITE), _
             Security, _
             OPEN_EXISTING, _
             0&, _
            0)
'Debug.Print "write handle"; WriteHandle
If WriteHandle = 0 Then
  OpenPipe = 0
  Exit Function
End If

  
ReadPathName = Left(DevicePathName, Len(DevicePathName) - 2) & "\PIPE1"
 ReadHandle = CreateFile _
            (ReadPathName, _
            GENERIC_READ Or GENERIC_WRITE, _
            (FILE_SHARE_READ Or FILE_SHARE_WRITE), _
            Security, _
            OPEN_EXISTING, _
            0&, _
            0)
If ReadHandle = 0 Then
  OpenPipe = 0
  Exit Function
End If

OpenPipe = 1
End Function

Public Function CBWTest_New_no_card(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

If PreSlotStatus <> 1 Then
    CBWTest_New_no_card = 4
    Exit Function
End If
CBWDataTransferLength = 1024
CBWTest_New_no_card = 0
If LBA > 25 * 1024 Then
LBA = 0
End If

 TmpString = ""
If ReaderExist = 0 Then
Do
DoEvents
 
               Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
              TmpString = GetDeviceName(Vid_PID)
  
Loop While TmpString = "" And TimerCounter < 10
End If
If ReaderExist = 0 And TmpString <> "" Then
  ReaderExist = 1
End If

If ReaderExist = 0 And TmpString = "" Then
  CBWTest_New_no_card = 0   ' no readerExist
  ReaderExist = 0
  Exit Function
End If

 


If OpenPipe = 0 Then
  CBWTest_New_no_card = 2   ' Write fail
  Exit Function
End If


'  TmpInteger = TestUnitSpeed(Lun)
    
 '   If TmpInteger = 0 Then
        
 '      CBWTest_New_no_card = 2   ' usb 2.0 high speed fail
 '      UsbSpeedTestResult = 2
 '      Exit Function
 '   End If
 '   TmpInteger = 0

 


TmpInteger = TestUnitReady(Lun)
If TmpInteger = 0 Then
    TmpInteger = RequestSense(Lun)
    
    
    If TmpInteger = 0 Or RequestSenseData(12) <> 58 Then
    
       CBWTest_New_no_card = 2  'Write fail
       Exit Function
    End If
    
End If





TmpInteger = TestUnitReady(Lun)

If TmpInteger = 0 Then
    TmpInteger = RequestSense(Lun)
    
    If TmpInteger = 0 Or RequestSenseData(12) <> 58 Then
    
       CBWTest_New_no_card = 2  'Write fail
       Exit Function
    End If
    
End If



CBWTest_New_no_card = 1
    

End Function
Public Function CBWTest_New_no_card_AU6350CF(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

If PreSlotStatus <> 1 Then
    CBWTest_New_no_card_AU6350CF = 4
    Exit Function
End If
CBWDataTransferLength = 1024
CBWTest_New_no_card_AU6350CF = 0
If LBA > 25 * 1024 Then
LBA = 0
End If

 TmpString = ""
If ReaderExist = 0 Then
Do
DoEvents
 
               Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
              TmpString = GetDeviceNameMulti(Vid_PID)
  
Loop While TmpString = "" And TimerCounter < 10
End If
If ReaderExist = 0 And TmpString <> "" Then
  ReaderExist = 1
End If

If ReaderExist = 0 And TmpString = "" Then
  CBWTest_New_no_card_AU6350CF = 0   ' no readerExist
  ReaderExist = 0
  Exit Function
End If

 


If OpenPipe = 0 Then
  CBWTest_New_no_card_AU6350CF = 2   ' Write fail
  Exit Function
End If


'  TmpInteger = TestUnitSpeed(Lun)
    
 '   If TmpInteger = 0 Then
        
 '      CBWTest_New_no_card_AU6350CF = 2   ' usb 2.0 high speed fail
 '      UsbSpeedTestResult = 2
 '      Exit Function
 '   End If
 '   TmpInteger = 0

 


TmpInteger = TestUnitReady(Lun)
If TmpInteger = 0 Then
    TmpInteger = RequestSense(Lun)
    
    
    If TmpInteger = 0 Or RequestSenseData(12) <> 58 Then
    
       CBWTest_New_no_card_AU6350CF = 2  'Write fail
       Exit Function
    End If
    
End If





TmpInteger = TestUnitReady(Lun)

If TmpInteger = 0 Then
    TmpInteger = RequestSense(Lun)
    
    If TmpInteger = 0 Or RequestSenseData(12) <> 58 Then
    
       CBWTest_New_no_card_AU6350CF = 2  'Write fail
       Exit Function
    End If
    
End If



CBWTest_New_no_card_AU6350CF = 1
    

End Function
Public Function CBWTest_New_no_card_AU6352DF(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

If PreSlotStatus <> 1 Then
    CBWTest_New_no_card_AU6352DF = 4
    Exit Function
End If
CBWDataTransferLength = 1024
CBWTest_New_no_card_AU6352DF = 0
If LBA > 25 * 1024 Then
LBA = 0
End If

 TmpString = ""
If ReaderExist = 0 Then
Do
DoEvents
 
               Call MsecDelay(0.1)
            TimerCounter = TimerCounter + 1
              TmpString = GetDeviceNameMulti(Vid_PID)
  
Loop While TmpString = "" And TimerCounter < 10
End If
If ReaderExist = 0 And TmpString <> "" Then
  ReaderExist = 1
End If

If ReaderExist = 0 And TmpString = "" Then
  CBWTest_New_no_card_AU6352DF = 0   ' no readerExist
  ReaderExist = 0
  Exit Function
End If

If OpenPipe = 0 Then
  CBWTest_New_no_card_AU6352DF = 2   ' Write fail
  Exit Function
End If


TmpInteger = TestUnitReady(Lun)
If TmpInteger = 0 Then
    TmpInteger = RequestSense(Lun)
    
    
    If TmpInteger = 0 Or RequestSenseData(12) <> 40 Then
    
       CBWTest_New_no_card_AU6352DF = 2  'Write fail
       Exit Function
    End If
    
End If

TmpInteger = TestUnitReady(Lun)

If TmpInteger = 0 Then
    TmpInteger = RequestSense(Lun)
    
    If TmpInteger = 0 Or RequestSenseData(12) <> 40 Then
    
       CBWTest_New_no_card_AU6352DF = 2  'Write fail
       Exit Function
    End If
    
End If



CBWTest_New_no_card_AU6352DF = 1
    

End Function

Public Function CBWTest_New_AU6366LFF21(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

   CBWDataTransferLength = 2048
 
'   For i = 0 To CBWDataTransferLength - 1
    
'         ReadData(i) = 0

'   Next

    If PreSlotStatus <> 1 Then
        CBWTest_New_AU6366LFF21 = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New_AU6366LFF21 = 0
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
      CBWTest_New_AU6366LFF21 = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_New_AU6366LFF21 = 2   ' Write fail
      Exit Function
    End If
 
    '======================================
    
    
     ' for unitSpeed
    
     TmpInteger = TestUnitSpeed(Lun)
    
     If TmpInteger = 0 Then
        
        CBWTest_New_AU6366LFF21 = 2   ' usb 2.0 high speed fail
        UsbSpeedTestResult = 2
        Exit Function
     End If
    
    
 
    TmpInteger = TestUnitReady(Lun)
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
           CBWTest_New_AU6366LFF21 = 2  'Write fail
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
         CBWTest_New_AU6366LFF21 = 2  'write fail
        '  Exit Function
     End If
    
      
    TmpInteger = Write_Data_LED(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_AU6366LFF21 = 2  'write fail
        Exit Function
    End If
    
    TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_AU6366LFF21 = 3    'Read fail
        Exit Function
    End If
     
    For i = 0 To CBWDataTransferLength - 1
    
        If ReadData(i) <> Pattern(i) Then
          CBWTest_New_AU6366LFF21 = 3    'Read fail
          Exit Function
        End If
    
    Next
    
    If Left(ChipName, 10) = "AU6371DLF2" Then
    
      TmpInteger = Read_CapacityAU6371(LBA, Lun, 8)
     If TmpInteger = 0 Then
      CBWTest_New_AU6366LFF21 = 3
      Exit Function
     End If
      
    End If
    
    
    
    CBWTest_New_AU6366LFF21 = 1
        
    
    End Function
Public Function CBWTest_New(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long
    
    CBWDataTransferLength = 4096
   For i = 0 To CBWDataTransferLength - 1
    
         ReadData(i) = 0

   Next

    If PreSlotStatus <> 1 Then
        CBWTest_New = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New = 0
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
      CBWTest_New = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_New = 2   ' Write fail
      Exit Function
    End If
 
    '======================================
    
    
     ' for unitSpeed
    
     TmpInteger = TestUnitSpeed(Lun)
    
     If TmpInteger = 0 Then
        
        CBWTest_New = 2   ' usb 2.0 high speed fail
        UsbSpeedTestResult = 2
        Exit Function
     End If
      
 
    TmpInteger = TestUnitReady(Lun)
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
           CBWTest_New = 2  'Write fail
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
         CBWTest_New = 2  'write fail
        '  Exit Function
     End If
      
    TmpInteger = Write_Data(LBA, Lun, CBWDataTransferLength)
    
    If TmpInteger = 0 Then
        Tester.Print "write fail"
        CBWTest_New = 2  'write fail
        Exit Function
    End If
    
    TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
     
'    If ChipName = "AU6479ULF23" Then
'        If TmpInteger = 0 Then
'            CBWTest_New = 3    'Read fail
'            For i = 0 To 3
'                Tester.Print "read 4k fail-" & (i + 1)
'                TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
'                If TmpInteger <> 0 Then
'                    Exit For
'                End If
'            Next
'
'            If TmpInteger = 0 And i = 4 Then
'                Exit Function
'            End If
'        End If
'    Else
        If TmpInteger = 0 Then
            CBWTest_New = 3    'Read fail
            Exit Function
        End If
'    End If
     
    For i = 0 To CBWDataTransferLength - 1
    
        If ReadData(i) <> Pattern(i) Then
          CBWTest_New = 3    'Read fail
          Exit Function
        End If
    
    Next
    
    
    If Left(ChipName, 10) = "AU6371DLF2" Then
    
      TmpInteger = Read_CapacityAU6371(LBA, Lun, 8)
     If TmpInteger = 0 Then
      CBWTest_New = 3
      Exit Function
     End If
      
     End If
    
    
    
    CBWTest_New = 1
        
    
    End Function
Public Function CBWTest_New_6336(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i, j As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long
    
    CBWDataTransferLength = 4096
   For i = 0 To CBWDataTransferLength - 1
    
         ReadData(i) = 0

   Next

    If PreSlotStatus <> 1 Then
        CBWTest_New_6336 = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New_6336 = 0
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
      CBWTest_New_6336 = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_New_6336 = 2   ' Write fail
      Exit Function
    End If
 
    '======================================
    
    
     ' for unitSpeed
     TmpInteger = TestUnitSpeed(Lun)
    
     If TmpInteger = 0 Then
        
        CBWTest_New_6336 = 2   ' usb 2.0 high speed fail
        UsbSpeedTestResult = 2
        Exit Function
     End If
      
    For j = 1 To 5
        TmpInteger = TestUnitReady(Lun)
        If TmpInteger = 1 Then
            Exit For
        End If
    Next
 
    'TmpInteger = TestUnitReady(Lun)
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
           CBWTest_New_6336 = 2  'Write fail
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
         CBWTest_New_6336 = 2  'write fail
        '  Exit Function
     End If
      
    TmpInteger = Write_Data(LBA, Lun, CBWDataTransferLength)
    
    If TmpInteger = 0 Then
        Tester.Print "write fail"
        CBWTest_New_6336 = 2  'write fail
        Exit Function
    End If
    
    TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
     
'    If ChipName = "AU6479ULF23" Then
'        If TmpInteger = 0 Then
'            CBWTest_New = 3    'Read fail
'            For i = 0 To 3
'                Tester.Print "read 4k fail-" & (i + 1)
'                TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
'                If TmpInteger <> 0 Then
'                    Exit For
'                End If
'            Next
'
'            If TmpInteger = 0 And i = 4 Then
'                Exit Function
'            End If
'        End If
'    Else
        If TmpInteger = 0 Then
            CBWTest_New_6336 = 3    'Read fail
            Exit Function
        End If
'    End If
     
    For i = 0 To CBWDataTransferLength - 1
    
        If ReadData(i) <> Pattern(i) Then
          CBWTest_New_6336 = 3    'Read fail
          Exit Function
        End If
    
    Next
    
    
    If Left(ChipName, 10) = "AU6371DLF2" Then
    
      TmpInteger = Read_CapacityAU6371(LBA, Lun, 8)
     If TmpInteger = 0 Then
      CBWTest_New_6336 = 3
      Exit Function
     End If
      
     End If
    
    
    
    CBWTest_New_6336 = 1
        
    
    End Function
Public Function CBWTest_Simple(Lun As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer


TmpInteger = WaitDevOn(Vid_PID)
If TmpInteger = 0 Then
    CBWTest_Simple = 0   ' usb 2.0 high speed fail
    Exit Function
End If


'=======================================
If OpenPipe = 0 Then
    CBWTest_Simple = 2   ' Write fail
    ClosePipe
    Exit Function
End If
 
'======================================
    
    
' for unitSpeed
TmpInteger = TestUnitSpeed(Lun)

If TmpInteger = 0 Then

    CBWTest_Simple = 2   ' usb 2.0 high speed fail
    UsbSpeedTestResult = 2
    ClosePipe
    Exit Function
End If
    
     
 
TmpInteger = TestUnitReady(Lun)
If TmpInteger = 0 Then
    TmpInteger = RequestSense(Lun)
    
    If TmpInteger = 0 Then
    
        CBWTest_Simple = 2  'Write fail
        ClosePipe
        Exit Function
    End If
    
End If

TmpInteger = Read_Data1(LBA, Lun, 512)
TmpInteger = Read_Data1(LBA, Lun, 512)

If TmpInteger = 0 Then
    CBWTest_Simple = 2  'write fail
    ClosePipe
    Exit Function
End If
    
'TmpInteger = Read_CapacityAU6371(LBA, Lun, 8)
'If TmpInteger = 0 Then
'    CBWTest_Simple = 3
'    ClosePipe
'    Exit Function
'End If
    
ClosePipe
ClosePipe
CBWTest_Simple = 1
        
    
End Function

Public Function CBWTest_New_PipeReady(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

CBWDataTransferLength = 4096

If PreSlotStatus <> 1 Then
    CBWTest_New_PipeReady = 4
    Exit Function
End If

For i = 0 To CBWDataTransferLength - 1
      ReadData(i) = 0
Next
    
'========================================
CBWTest_New_PipeReady = 0
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
    CBWTest_New_PipeReady = 0   ' no readerExist
    ReaderExist = 0
    Exit Function
End If
'=======================================

' for unitSpeed

TmpInteger = TestUnitSpeed(Lun)

If TmpInteger = 0 Then
    CBWTest_New_PipeReady = 2   ' usb 2.0 high speed fail
    UsbSpeedTestResult = 2
    Exit Function
End If
     
TmpInteger = TestUnitReady(Lun)
If TmpInteger = 0 Then
    TmpInteger = RequestSense(Lun)
    If TmpInteger = 0 Then
        CBWTest_New_PipeReady = 2  'Write fail
        Exit Function
    End If
End If

TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
     
   
If TmpInteger = 0 Then
    CBWTest_New_PipeReady = 2  'write fail
    'Exit Function
End If
      
TmpInteger = Write_Data(LBA, Lun, CBWDataTransferLength)

If TmpInteger = 0 Then
    CBWTest_New_PipeReady = 2  'write fail
    Exit Function
End If
    
TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
 
If TmpInteger = 0 Then
    CBWTest_New_PipeReady = 3    'Read fail
    Exit Function
End If
     
For i = 0 To CBWDataTransferLength - 1
    If ReadData(i) <> Pattern(i) Then
        CBWTest_New_PipeReady = 3    'Read fail
        Exit Function
    End If
Next
    
CBWTest_New_PipeReady = 1
        
    
End Function
    
Public Function CBWTest_New_MultiDevice(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String, MemNo As Integer) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

    CBWDataTransferLength = 4096
'   For i = 0 To CBWDataTransferLength - 1
    
'         ReadData(i) = 0

'   Next
    
    ReaderExist = 0
    HubPort = MemNo
    
    If PreSlotStatus <> 1 Then
        CBWTest_New_MultiDevice = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New_MultiDevice = 0
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
      CBWTest_New_MultiDevice = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_New_MultiDevice = 2   ' Write fail
      Exit Function
    End If
 
    '======================================
    
    
     ' for unitSpeed
    
     TmpInteger = TestUnitSpeed(Lun)
    
     If TmpInteger = 0 Then
        
        CBWTest_New_MultiDevice = 2   ' usb 2.0 high speed fail
        UsbSpeedTestResult = 2
        Exit Function
     End If
    
    
 
    TmpInteger = TestUnitReady(Lun)
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
           CBWTest_New_MultiDevice = 2  'Write fail
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
         CBWTest_New_MultiDevice = 2  'write fail
        '  Exit Function
     End If
      
    TmpInteger = Write_Data(LBA, Lun, CBWDataTransferLength)
    
    If TmpInteger = 0 Then
        CBWTest_New_MultiDevice = 2  'write fail
        Exit Function
    End If
    
    TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_MultiDevice = 3    'Read fail
        Exit Function
    End If
     
    For i = 0 To CBWDataTransferLength - 1
    
        If ReadData(i) <> Pattern(i) Then
          CBWTest_New_MultiDevice = 3    'Read fail
          Exit Function
        End If
    
    Next
    
    
    CBWTest_New_MultiDevice = 1
        
    
End Function
Public Function AU6435_CBWTest_New(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte

Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

   CBWDataTransferLength = 2048
 
   For i = 0 To CBWDataTransferLength - 1
    
         ReadData(i) = 0

   Next

    If PreSlotStatus <> 1 Then
        AU6435_CBWTest_New = 4
        Exit Function
    End If
    '========================================
   
    AU6435_CBWTest_New = 0
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
      AU6435_CBWTest_New = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    
    If OpenPipe = 0 Then
      AU6435_CBWTest_New = 2   ' Write fail
      Exit Function
    End If
 
    '======================================
    
    
     ' for unitSpeed
    
     TmpInteger = TestUnitSpeed(Lun)
    
     If TmpInteger = 0 Then
        
        AU6435_CBWTest_New = 2   ' usb 2.0 high speed fail
        UsbSpeedTestResult = 2
        Exit Function
     End If
 
    TmpInteger = TestUnitReady(Lun)
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        If TmpInteger = 0 Then
            Call MsecDelay(0.02)
            TmpInteger = RequestSense(Lun)
        End If
        If TmpInteger = 0 Then
           AU6435_CBWTest_New = 2  'Write fail
           Exit Function
        End If
        
    End If
    '======================================
  '  If ChipName = "AU6371" Or ChipName = "AU6371S3" Then
  '      TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
  '  End If
    
    
     TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
     
     If TmpInteger = 0 Then
        TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
     End If
     
     
    If TmpInteger = 0 Then
         AU6435_CBWTest_New = 2  'write fail
         'Exit Function
     End If
    
      
    TmpInteger = Write_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        TmpInteger = Write_Data(LBA, Lun, CBWDataTransferLength)
    End If
    
    If TmpInteger = 0 Then
        AU6435_CBWTest_New = 2  'write fail
        Exit Function
    End If
    
    TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
    End If
     
    If TmpInteger = 0 Then
        AU6435_CBWTest_New = 3    'Read fail
        Exit Function
    End If
     
    For i = 0 To CBWDataTransferLength - 1
    
        If ReadData(i) <> Pattern(i) Then
          AU6435_CBWTest_New = 3    'Read fail
          Exit Function
        End If
    
    Next
 
    AU6435_CBWTest_New = 1
        
    
    End Function

Public Function AU6990HW_Version() As Byte

    Dim i As Integer
    Dim WriteTest As Integer
    Dim TmpString As String
    Dim TmpInteger As Integer
    Dim TimerCounter As Integer
    Dim OldTimer
    Dim CBWDataTransferLength As Long
    Dim HWVersion As String
    Dim Vid_PID As String
        
        
    CBWDataTransferLength = 512
    Vid_PID = "vid"
    
    For i = 0 To CBWDataTransferLength - 1
        ReadData(i) = 0
    Next

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
      AU6990HW_Version = 0  ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      AU6990HW_Version = 4   ' Write fail
      Exit Function
    End If
 
    '======================================
    
    
     ' for unitSpeed
    
    TmpInteger = TestUnitSpeed(Lun)
    
    If TmpInteger = 0 Then
        AU6990HW_Version = 2   ' usb 2.0 high speed fail
        UsbSpeedTestResult = 2
        Exit Function
    End If
    
    TmpInteger = TestUnitReady(Lun)
    
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
            AU6990HW_Version = 2  'Write fail
            Exit Function
        End If
        
    End If
    
        
    AU6990HW_Version = 1
        
    End Function
    
    Public Function AU6922HW_Version() As Byte

    Dim i As Integer
    Dim WriteTest As Integer
    Dim TmpString As String
    Dim TmpInteger As Integer
    Dim TimerCounter As Integer
    Dim OldTimer
    Dim CBWDataTransferLength As Long
    Dim HWVersion As String
    Dim Vid_PID As String
        
        
    CBWDataTransferLength = 512
    Vid_PID = "vid"
    
    For i = 0 To CBWDataTransferLength - 1
        ReadData(i) = 0
    Next

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
      AU6922HW_Version = 0  ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      AU6922HW_Version = 4   ' Write fail
      Exit Function
    End If
 
    '======================================
    
    
     ' for unitSpeed
    
    TmpInteger = TestUnitSpeed(Lun)
    
    If TmpInteger = 0 Then
        AU6922HW_Version = 2   ' usb 2.0 high speed fail
        UsbSpeedTestResult = 2
        Exit Function
    End If
    
    TmpInteger = TestUnitReady(Lun)
    
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
            AU6922HW_Version = 2  'Write fail
            Exit Function
        End If
        
    End If
    
        
    AU6922HW_Version = 1
        
End Function

Public Function CBWTest_NewCFOverCurrentRWnoPower(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

   CBWDataTransferLength = 2048
 
'   For i = 0 To CBWDataTransferLength - 1
    
'         ReadData(i) = 0

'   Next

    If PreSlotStatus <> 1 Then
        CBWTest_NewCFOverCurrentRWnoPower = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_NewCFOverCurrentRWnoPower = 0
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
      CBWTest_NewCFOverCurrentRWnoPower = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_NewCFOverCurrentRWnoPower = 2   ' Write fail
      Exit Function
    End If
 
    '======================================
    
    
     ' for unitSpeed
    
     TmpInteger = TestUnitSpeed(Lun)
    
     If TmpInteger = 0 Then
        
        CBWTest_NewCFOverCurrentRWnoPower = 2   ' usb 2.0 high speed fail
        UsbSpeedTestResult = 2
        Exit Function
     End If
    
    
       TmpInteger = SetOverCurrent(Lun)
    
    If TmpInteger <> 1 Then
       CBWTest_NewCFOverCurrentRWnoPower = 2
       Exit Function
    End If
       
   
    TmpInteger = TestUnitReady(Lun)
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
           CBWTest_NewCFOverCurrentRWnoPower = 2  'Write fail
           Exit Function
        End If
        
    End If
    
   
     CBWTest_NewCFOverCurrentRWnoPower = Read_OverCurrent(0, 0, 64)
                      
    
    If CBWTest_NewCFOverCurrentRWnoPower = 0 Then
    CBWTest_NewCFOverCurrentRWnoPower = 2
    Exit Function
    End If
    
      If CBWTest_NewCFOverCurrentRWnoPower = 3 Then
    CBWTest_NewCFOverCurrentRWnoPower = 5
    Exit Function
    End If
    
    
    '======================================
  '  If ChipName = "AU6371" Or ChipName = "AU6371S3" Then
  '      TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
  '  End If
   
     TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
   
     TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
   
    If TmpInteger = 0 Then
         CBWTest_NewCFOverCurrentRWnoPower = 2  'write fail
        '  Exit Function
     End If
    
      
    TmpInteger = Write_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_NewCFOverCurrentRWnoPower = 2  'write fail
        Exit Function
    End If
    
    TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_NewCFOverCurrentRWnoPower = 3    'Read fail
        Exit Function
    End If
     
    For i = 0 To CBWDataTransferLength - 1
    
        If ReadData(i) <> Pattern(i) Then
          CBWTest_NewCFOverCurrentRWnoPower = 3    'Read fail
          Exit Function
        End If
    
    Next
    
    If Left(ChipName, 10) = "AU6371DLF2" Then
    
      TmpInteger = Read_CapacityAU6371(LBA, Lun, 8)
     If TmpInteger = 0 Then
      CBWTest_NewCFOverCurrentRWnoPower = 3
      Exit Function
     End If
      
    End If
    
    
    
    CBWTest_NewCFOverCurrentRWnoPower = 1
        
    
    End Function
    
Public Function CBWTest_NewCFOverCurrentRW(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

   CBWDataTransferLength = 2048
 
'   For i = 0 To CBWDataTransferLength - 1
    
'         ReadData(i) = 0

'   Next

    If PreSlotStatus <> 1 Then
        CBWTest_NewCFOverCurrentRW = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_NewCFOverCurrentRW = 0
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
      CBWTest_NewCFOverCurrentRW = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_NewCFOverCurrentRW = 2   ' Write fail
      Exit Function
    End If
 
    '======================================
    
    
     ' for unitSpeed
    
     TmpInteger = TestUnitSpeed(Lun)
    
     If TmpInteger = 0 Then
        
        CBWTest_NewCFOverCurrentRW = 2   ' usb 2.0 high speed fail
        UsbSpeedTestResult = 2
        Exit Function
     End If
    
    
       TmpInteger = SetOverCurrent(Lun)
    
    If TmpInteger <> 1 Then
       CBWTest_NewCFOverCurrentRW = 2
       Exit Function
    End If
       
    Call PowerSet2(1, "3.20", "0.5", 1, "3.2", "0.5", 1)
      Call MsecDelay(1.5)
    TmpInteger = TestUnitReady(Lun)
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
           CBWTest_NewCFOverCurrentRW = 2  'Write fail
           Exit Function
        End If
        
    End If
    
   
     CBWTest_NewCFOverCurrentRW = Read_OverCurrent(0, 0, 64)
                      
    
    If CBWTest_NewCFOverCurrentRW = 0 Then
    CBWTest_NewCFOverCurrentRW = 2
    Exit Function
    End If
    
      If CBWTest_NewCFOverCurrentRW = 3 Then
    CBWTest_NewCFOverCurrentRW = 5
    Exit Function
    End If
    
       Call PowerSet2(1, "3.25", "0.5", 1, "3.2", "0.5", 1)
     Call MsecDelay(1#)
    '======================================
  '  If ChipName = "AU6371" Or ChipName = "AU6371S3" Then
  '      TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
  '  End If
   
     TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
   
     TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
   
    If TmpInteger = 0 Then
         CBWTest_NewCFOverCurrentRW = 2  'write fail
        '  Exit Function
     End If
    
      
    TmpInteger = Write_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_NewCFOverCurrentRW = 2  'write fail
        Exit Function
    End If
    
    TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_NewCFOverCurrentRW = 3    'Read fail
        Exit Function
    End If
     
    For i = 0 To CBWDataTransferLength - 1
    
        If ReadData(i) <> Pattern(i) Then
          CBWTest_NewCFOverCurrentRW = 3    'Read fail
          Exit Function
        End If
    
    Next
    
    If Left(ChipName, 10) = "AU6371DLF2" Then
    
      TmpInteger = Read_CapacityAU6371(LBA, Lun, 8)
     If TmpInteger = 0 Then
      CBWTest_NewCFOverCurrentRW = 3
      Exit Function
     End If
      
    End If
    
    
    
    CBWTest_NewCFOverCurrentRW = 1
        
    
    End Function
    
Public Function CBWTest_NewOverCurrent5(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

   CBWDataTransferLength = 2048
 
'   For i = 0 To CBWDataTransferLength - 1
    
'         ReadData(i) = 0

'   Next

    If PreSlotStatus <> 1 Then
        CBWTest_NewOverCurrent5 = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_NewOverCurrent5 = 0
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
      CBWTest_NewOverCurrent5 = 0    ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_NewOverCurrent5 = 2    ' Write fail
      Exit Function
    End If
 
    '======================================
    
    
     ' for unitSpeed
    
     TmpInteger = TestUnitSpeed(Lun)
    
     If TmpInteger = 0 Then
        
        CBWTest_NewOverCurrent5 = 2  ' usb 2.0 high speed fail
        UsbSpeedTestResult = 2
        Exit Function
     End If
    
    
'      TmpInteger = SetOverCurrent5(Lun)
    
  '   If TmpInteger <> 1 Then
      '  CBWTest_NewOverCurrent5 = 2
    '    Exit Function
  '   End If
       
 
    TmpInteger = TestUnitReady5(Lun)
    
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
            CBWTest_NewOverCurrent5 = 2  'Write fail
           Exit Function
        End If
        
    End If
    '======================================

       
                   
    CBWTest_NewOverCurrent5 = Read_OverCurrent(0, 0, 64)
                      
    
    If CBWTest_NewOverCurrent5 = 0 Then
    CBWTest_NewOverCurrent5 = 2
    End If
   
        
    
    End Function
Public Function CBWTest_NewOverCurrent(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

   CBWDataTransferLength = 2048
 
'   For i = 0 To CBWDataTransferLength - 1
    
'         ReadData(i) = 0

'   Next

    If PreSlotStatus <> 1 Then
        CBWTest_NewOverCurrent = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_NewOverCurrent = 0
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
      CBWTest_NewOverCurrent = 0    ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_NewOverCurrent = 2    ' Write fail
      Exit Function
    End If
 
    '======================================
    
    
     ' for unitSpeed
    
     TmpInteger = TestUnitSpeed(Lun)
    
     If TmpInteger = 0 Then
        
        CBWTest_NewOverCurrent = 2  ' usb 2.0 high speed fail
        UsbSpeedTestResult = 2
        Exit Function
     End If
    
    
     TmpInteger = SetOverCurrent(Lun)
    
    If TmpInteger <> 1 Then
       CBWTest_NewOverCurrent = 2
       Exit Function
    End If
       
 
    TmpInteger = TestUnitReady(Lun)
    
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
            CBWTest_NewOverCurrent = 2  'Write fail
           Exit Function
        End If
        
    End If
    '======================================

       
                   
    CBWTest_NewOverCurrent = Read_OverCurrent(0, 0, 64)
                      
    
    If CBWTest_NewOverCurrent = 0 Then
    CBWTest_NewOverCurrent = 2
    End If
   
        
    
    End Function
    
    Public Function CBWTest_NewOverCurrent2(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

   CBWDataTransferLength = 2048
 
'   For i = 0 To CBWDataTransferLength - 1
    
'         ReadData(i) = 0

'   Next

    If PreSlotStatus <> 1 Then
        CBWTest_NewOverCurrent2 = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_NewOverCurrent2 = 0
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
      CBWTest_NewOverCurrent2 = 0    ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_NewOverCurrent2 = 2    ' Write fail
      Exit Function
    End If
 
    '======================================
    
    
     ' for unitSpeed
    
      TmpInteger = TestUnitSpeed(Lun)
    
      If TmpInteger = 0 Then
        
        CBWTest_NewOverCurrent2 = 2  ' usb 2.0 high speed fail
         UsbSpeedTestResult = 2
         Exit Function
      End If
    
    
    ' TmpInteger = SetOverCurrent(Lun)
    
   ' If TmpInteger <> 1 Then
   '    CBWTest_NewOverCurrent2 = 2
   '    Exit Function
   ' End If
       
    
    TmpInteger = TestUnitReadyReadOverCurrent(Lun)
    
   ' If TmpInteger = 0 Then
   '     TmpInteger = RequestSense(Lun)
        
   '     If TmpInteger = 0 Then
        
   '         CBWTest_NewOverCurrent2 = 2  'Write fail
   '        Exit Function
   '     End If
        
   ' End If
    '======================================

       
                   
   ' CBWTest_NewOverCurrent2 = Read_OverCurrent(0, 0, 64)
                      
    
     
    CBWTest_NewOverCurrent2 = TmpInteger
    
   
        
    
    End Function
Public Function CBWTest_NewOverCurrent3(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

   CBWDataTransferLength = 2048
 
'   For i = 0 To CBWDataTransferLength - 1
    
'         ReadData(i) = 0

'   Next

    If PreSlotStatus <> 1 Then
        CBWTest_NewOverCurrent3 = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_NewOverCurrent3 = 0
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
      CBWTest_NewOverCurrent3 = 0    ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_NewOverCurrent3 = 2    ' Write fail
      Exit Function
    End If
 
    '======================================
    
    
     ' for unitSpeed
    
      TmpInteger = TestUnitSpeed(Lun)
    
      If TmpInteger = 0 Then
        
        CBWTest_NewOverCurrent3 = 2  ' usb 2.0 high speed fail
         UsbSpeedTestResult = 2
         Exit Function
      End If
    
    
    ' TmpInteger = SetOverCurrent(Lun)
    
   ' If TmpInteger <> 1 Then
   '    CBWTest_NewOverCurrent3 = 2
   '    Exit Function
   ' End If
       
    
    TmpInteger = TestUnitReadyReadOverCurrentA2D(Lun)
    
 
       
                   
  
     
    CBWTest_NewOverCurrent3 = TmpInteger
    
   
        
    
    End Function
    
Public Function CBWTest_NewOverCurrent4(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

   CBWDataTransferLength = 2048
 
'   For i = 0 To CBWDataTransferLength - 1
    
'         ReadData(i) = 0

'   Next

    If PreSlotStatus <> 1 Then
        CBWTest_NewOverCurrent4 = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_NewOverCurrent4 = 0
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
      CBWTest_NewOverCurrent4 = 0    ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_NewOverCurrent4 = 2    ' Write fail
      Exit Function
    End If
 
    '======================================
    
    
     ' for unitSpeed
    
      TmpInteger = TestUnitSpeed(Lun)
    
      If TmpInteger = 0 Then
        
        CBWTest_NewOverCurrent4 = 2  ' usb 2.0 high speed fail
         UsbSpeedTestResult = 2
         Exit Function
      End If
    
    
 '    TmpInteger = SetOverCurrent5(Lun)
    
   ' If TmpInteger <> 1 Then
    '  CBWTest_NewOverCurrent4 = 2
   '    Exit Function
  '  End If
       
    
    TmpInteger = TestUnitReadyReadOverCurrentA2D4(Lun)
    
 
       
                   
  
     
    CBWTest_NewOverCurrent4 = TmpInteger
    
   
        
    
    End Function
    
Public Function CBWTest_New_AU6371Fail(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Long
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

   CBWDataTransferLength = 65536
 
'   For i = 0 To CBWDataTransferLength - 1
    
'         ReadData(i) = 0

'   Next

    If PreSlotStatus <> 1 Then
        CBWTest_New_AU6371Fail = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New_AU6371Fail = 0
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
      CBWTest_New_AU6371Fail = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_New_AU6371Fail = 2   ' Write fail
      Exit Function
    End If
 
    '======================================
    
    
     ' for unitSpeed
    
     TmpInteger = TestUnitSpeed(Lun)
    
     If TmpInteger = 0 Then
        
        CBWTest_New_AU6371Fail = 2   ' usb 2.0 high speed fail
        UsbSpeedTestResult = 2
        Exit Function
     End If
    
    
 
    TmpInteger = TestUnitReady(Lun)
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
           CBWTest_New_AU6371Fail = 2  'Write fail
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
         CBWTest_New_AU6371Fail = 2  'write fail
        '  Exit Function
     End If
    
      
    TmpInteger = Write_Data_AU6371(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_AU6371Fail = 2  'write fail
        Exit Function
    End If
    
    TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_AU6371Fail = 3    'Read fail
        Exit Function
    End If
     
    For i = 0 To CBWDataTransferLength - 1
    
        If ReadData(i) <> AU6371Pattern(i) Then
          CBWTest_New_AU6371Fail = 3    'Read fail
          Exit Function
        End If
    
    Next
    
    
    
    
    
    CBWTest_New_AU6371Fail = 1
        
    
    End Function
    
Public Function CBWTest_New_AU6371DLS41(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

  Dim Sv1 As String
          Dim Sv2 As String
              '  Sv1 = "2.9"
               ' Sv2 = "1.40"
               
                  Sv1 = "3.3"
                 Sv2 = "1.4"
                 ' Call PowerSet2(2, "3.3", "0.5", 1, "1.77", "0.5", 1)
                 '  Call MsecDelay(0.4)
                    Call PowerSet2(0, Sv1, "0.5", 1, Sv2, "0.5", 1)
                    Call MsecDelay(2.3)

   CBWDataTransferLength = 2048
 
'   For i = 0 To CBWDataTransferLength - 1
    
'         ReadData(i) = 0

'   Next

    If PreSlotStatus <> 1 Then
        CBWTest_New_AU6371DLS41 = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New_AU6371DLS41 = 0
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
      CBWTest_New_AU6371DLS41 = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_New_AU6371DLS41 = 2   ' Write fail
      Exit Function
    End If
 
    '======================================
    
    
     ' for unitSpeed
    
     TmpInteger = TestUnitSpeed(Lun)
    
     If TmpInteger = 0 Then
        
        CBWTest_New_AU6371DLS41 = 2   ' usb 2.0 high speed fail
        UsbSpeedTestResult = 2
        Exit Function
     End If
     
     
        Sv1 = "2.88"
    
                '  Call PowerSet2(2, Sv1, "0.5", 1, Sv2, "0.5", 1)
                 '  Call MsecDelay(0.4)
                    Call PowerSet2(1, Sv1, "0.5", 1, Sv2, "0.5", 1)
                    Call MsecDelay(1.8)
            
 
    TmpInteger = TestUnitReady(Lun)
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
           CBWTest_New_AU6371DLS41 = 2  'Write fail
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
         CBWTest_New_AU6371DLS41 = 2  'write fail
        '  Exit Function
     End If
    
      
    TmpInteger = Write_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_AU6371DLS41 = 2  'write fail
        Exit Function
    End If
    
    TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_AU6371DLS41 = 3    'Read fail
        Exit Function
    End If
     
    For i = 0 To CBWDataTransferLength - 1
    
        If ReadData(i) <> Pattern(i) Then
          CBWTest_New_AU6371DLS41 = 3    'Read fail
          Exit Function
        End If
    
    Next
   
    
    
    
    CBWTest_New_AU6371DLS41 = 1
        
    
    End Function
Public Function CBWTest_New_SD_Speed(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String, BitWidth As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

   CBWDataTransferLength = 2048
 
'   For i = 0 To CBWDataTransferLength - 1
    
'         ReadData(i) = 0

'   Next

    If PreSlotStatus <> 1 Then
        CBWTest_New_SD_Speed = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New_SD_Speed = 0
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
      CBWTest_New_SD_Speed = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_New_SD_Speed = 2   ' Write fail
      Exit Function
    End If
 
    '======================================
    
    
     ' for unitSpeed
    
     TmpInteger = TestUnitSpeed(Lun)
    
     If TmpInteger = 0 Then
        
        CBWTest_New_SD_Speed = 2   ' usb 2.0 high speed fail
        UsbSpeedTestResult = 2
        Exit Function
     End If
    
    
 
    TmpInteger = TestUnitReady(Lun)
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
           CBWTest_New_SD_Speed = 2  'Write fail
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
         CBWTest_New_SD_Speed = 2  'write fail
        '  Exit Function
     End If
    
     TmpInteger = Read_SD_Speed(LBA, Lun, 64, BitWidth)
     
      
    TmpInteger = Write_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_SD_Speed = 2  'write fail
        Exit Function
    End If
    
    TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_SD_Speed = 3    'Read fail
        Exit Function
    End If
     
    For i = 0 To CBWDataTransferLength - 1
    
        If ReadData(i) <> Pattern(i) Then
          CBWTest_New_SD_Speed = 3    'Read fail
          Exit Function
        End If
    
    Next
    
    If Left(ChipName, 10) = "AU6371DLF2" Then
    
      TmpInteger = Read_CapacityAU6371(LBA, Lun, 8)
     If TmpInteger = 0 Then
      CBWTest_New_SD_Speed = 3
      Exit Function
     End If
      
    End If
    
    
    
    CBWTest_New_SD_Speed = 1
        
    
    End Function
Public Function CBWTest_NewAU6336ZFF20(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

   CBWDataTransferLength = 2048
 
'   For i = 0 To CBWDataTransferLength - 1
    
'         ReadData(i) = 0

'   Next

    If PreSlotStatus <> 1 Then
        CBWTest_NewAU6336ZFF20 = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_NewAU6336ZFF20 = 0
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
      CBWTest_NewAU6336ZFF20 = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_NewAU6336ZFF20 = 2   ' Write fail
      Exit Function
    End If
 
    '======================================
    
    
     ' for unitSpeed
    
   '  TmpInteger = TestUnitSpeed(Lun)
    
   '  If TmpInteger = 0 Then
        
   '     CBWTest_NewAU6336ZFF20 = 2   ' usb 2.0 high speed fail
   '     UsbSpeedTestResult = 2
   '     Exit Function
   '  End If
    
    
 
    TmpInteger = TestUnitReady(Lun)
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
           CBWTest_NewAU6336ZFF20 = 2  'Write fail
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
         CBWTest_NewAU6336ZFF20 = 2  'write fail
        '  Exit Function
     End If
    
      
    TmpInteger = Write_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_NewAU6336ZFF20 = 2  'write fail
        Exit Function
    End If
    
    TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_NewAU6336ZFF20 = 3    'Read fail
        Exit Function
    End If
     
    For i = 0 To CBWDataTransferLength - 1
    
        If ReadData(i) <> Pattern(i) Then
          CBWTest_NewAU6336ZFF20 = 3    'Read fail
          Exit Function
        End If
    
    Next
    
    If Left(ChipName, 10) = "AU6371DLF2" Then
    
      TmpInteger = Read_CapacityAU6371(LBA, Lun, 8)
     If TmpInteger = 0 Then
      CBWTest_NewAU6336ZFF20 = 3
      Exit Function
     End If
      
    End If
    
    
    
    CBWTest_NewAU6336ZFF20 = 1
        
    
    End Function
    
Public Function CBWTest_New_CIS(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

   CBWDataTransferLength = 2048
 
'   For i = 0 To CBWDataTransferLength - 1
    
'         ReadData(i) = 0

'   Next

    If PreSlotStatus <> 1 Then
        CBWTest_New_CIS = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New_CIS = 0
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
      CBWTest_New_CIS = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_New_CIS = 2   ' Write fail
      Exit Function
    End If
 
    '======================================
    
    
     ' for unitSpeed
    
     TmpInteger = TestUnitSpeed(Lun)
    
     If TmpInteger = 0 Then
        
        CBWTest_New_CIS = 2   ' usb 2.0 high speed fail
        UsbSpeedTestResult = 2
        Exit Function
     End If
    
    
    TmpInteger = ReInitial(0)
    TmpInteger = TestUnitReady(Lun)
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
           CBWTest_New_CIS = 2  'Write fail
           Exit Function
        End If
        
    End If
    '======================================
  '  If ChipName = "AU6371" Or ChipName = "AU6371S3" Then
  '      TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
  '  End If
  
    TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
   
     TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
     
     
      TmpInteger = Read_DataCIS(LBA, Lun, CBWDataTransferLength)
    
    If TmpInteger = 1 Then
         CBWTest_New_CIS = 2   'CIS Read Get 2048 byte Fail
           Exit Function
     End If
    
   
    
    
    CBWTest_New_CIS = 1
        
    
    End Function
    
Public Function CBWTest_New_NO_CIS(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

   CBWDataTransferLength = 2048
 
'   For i = 0 To CBWDataTransferLength - 1
    
'         ReadData(i) = 0

'   Next

    If PreSlotStatus <> 1 Then
        CBWTest_New_NO_CIS = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New_NO_CIS = 0
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
      CBWTest_New_NO_CIS = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_New_NO_CIS = 2   ' Write fail
      Exit Function
    End If
 
    '======================================
    
    
     ' for unitSpeed
    
     TmpInteger = TestUnitSpeed(Lun)
    
     If TmpInteger = 0 Then
        
        CBWTest_New_NO_CIS = 2   ' usb 2.0 high speed fail
        UsbSpeedTestResult = 2
        Exit Function
     End If
    
    
    TmpInteger = ReInitial(0)
    TmpInteger = TestUnitReady(Lun)
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
           CBWTest_New_NO_CIS = 2  'Write fail
           Exit Function
        End If
        
    End If
    '======================================
  '  If ChipName = "AU6371" Or ChipName = "AU6371S3" Then
  '      TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
  '  End If
  
    TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
   
     TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
     
     
      TmpInteger = Read_DataCIS(LBA, Lun, CBWDataTransferLength)
    
    If TmpInteger = 0 Then
         CBWTest_New_NO_CIS = 2   'CIS Read Get 2048 byte Fail
           Exit Function
     End If
    
   
    
    
    CBWTest_New_NO_CIS = 1
        
    
    End Function
Public Function CBWTest_New_NB(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

   CBWDataTransferLength = 2048
 
'   For i = 0 To CBWDataTransferLength - 1
    
'         ReadData(i) = 0

'   Next

    If PreSlotStatus <> 1 Then
        CBWTest_New_NB = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New_NB = 0
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
      CBWTest_New_NB = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_New_NB = 2   ' Write fail
      Exit Function
    End If
 
    '======================================
    
    
     ' for unitSpeed
    
     TmpInteger = TestUnitSpeed(Lun)
    
     If TmpInteger = 0 Then
        
        CBWTest_New_NB = 2   ' usb 2.0 high speed fail
        UsbSpeedTestResult = 2
        Exit Function
     End If
    
    
    TmpInteger = ReInitial(0)
    TmpInteger = TestUnitReady(Lun)
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
           CBWTest_New_NB = 2  'Write fail
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
         CBWTest_New_NB = 2  'write fail
        '  Exit Function
     End If
    
      
    TmpInteger = Write_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_NB = 2  'write fail
        Exit Function
    End If
    
    TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_NB = 3    'Read fail
        Exit Function
    End If
     
    For i = 0 To CBWDataTransferLength - 1
    
        If ReadData(i) <> Pattern(i) Then
          CBWTest_New_NB = 3    'Read fail
          Exit Function
        End If
    
    Next
    
    If Left(ChipName, 10) = "AU6371DLF2" Then
    
      TmpInteger = Read_CapacityAU6371(LBA, Lun, 8)
     If TmpInteger = 0 Then
      CBWTest_New_NB = 3
      Exit Function
     End If
      
    End If
    
    
    
    CBWTest_New_NB = 1
        
    
    End Function
    
    
Public Function CBWTest_NewAU6378AutoModeFail(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

   CBWDataTransferLength = 2048
 
'   For i = 0 To CBWDataTransferLength - 1
    
'         ReadData(i) = 0

'   Next

    If PreSlotStatus <> 1 Then
        CBWTest_NewAU6378AutoModeFail = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_NewAU6378AutoModeFail = 0
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
      CBWTest_NewAU6378AutoModeFail = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_NewAU6378AutoModeFail = 2   ' Write fail
      Exit Function
    End If
 
    '======================================
    
    
     ' for unitSpeed
    
     TmpInteger = TestUnitSpeed(Lun)
    
     If TmpInteger = 0 Then
        
        CBWTest_NewAU6378AutoModeFail = 2   ' usb 2.0 high speed fail
        UsbSpeedTestResult = 2
        Exit Function
     End If
    
    
    
    TmpInteger = TestUnitReady(Lun)
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
           CBWTest_NewAU6378AutoModeFail = 2  'Write fail
           Exit Function
        End If
        
    End If
    '======================================
  '  If ChipName = "AU6371" Or ChipName = "AU6371S3" Then
  '      TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
  '  End If
    
    TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
    
     CBWDataTransferLength = 4096  '8 sector
      TmpInteger = 0
     TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
      If TmpInteger = 0 Then
         CBWTest_NewAU6378AutoModeFail = 2  'write fail
        '  Exit Function
     End If
    
      CBWDataTransferLength = 65024  '127 sector
      TmpInteger = 0
     TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
      If TmpInteger = 0 Then
         CBWTest_NewAU6378AutoModeFail = 2  'write fail
        '  Exit Function
     End If
     
     
        
      CBWDataTransferLength = 65536  '128 sector
      TmpInteger = 0
     TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
      If TmpInteger = 0 Then
         CBWTest_NewAU6378AutoModeFail = 2  'write fail
        '  Exit Function
     End If
      
    
     
     CBWDataTransferLength = 2048
   
    
      
    TmpInteger = Write_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_NewAU6378AutoModeFail = 2  'write fail
        Exit Function
    End If
    
    TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_NewAU6378AutoModeFail = 3    'Read fail
        Exit Function
    End If
     
    For i = 0 To CBWDataTransferLength - 1
    
        If ReadData(i) <> Pattern(i) Then
          CBWTest_NewAU6378AutoModeFail = 3    'Read fail
          Exit Function
        End If
    
    Next
    
    If Left(ChipName, 10) = "AU6371DLF2" Then
    
      TmpInteger = Read_CapacityAU6371(LBA, Lun, 8)
     If TmpInteger = 0 Then
      CBWTest_NewAU6378AutoModeFail = 3
      Exit Function
     End If
      
    End If
    
    
    
    CBWTest_NewAU6378AutoModeFail = 1
        
    
    End Function
    
Public Function Read_CapacityAU6371(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte
Dim CBW(0 To 30) As Byte
Dim NumberOfBytesWritten As Long
Dim CBWDataTransferLen(0 To 3) As Byte
  
Dim TransferLen As Long
Dim TransferLenLSB As Byte
Dim TransferLenMSB As Byte
Dim i As Integer
Dim tmpV(0 To 2) As Long
Dim opcode As Byte

Dim CSW(0 To 12) As Byte

Dim NumberOfBytesRead As Long

Dim Capacity(0 To 7) As Byte

'Capacity(0) = &H0
'Capacity(1) = &H78
'Capacity(2) = &HFF
'Capacity(3) = &HFF
'Capacity(4) = &H0
'Capacity(5) = &H0
'Capacity(6) = &H2
'Capacity(7) = &H0

Capacity(0) = &H0
Capacity(1) = &H3
Capacity(2) = &HC9
Capacity(3) = &HFF
Capacity(4) = &H0
Capacity(5) = &H0
Capacity(6) = &H2
Capacity(7) = &H0

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

 
CBW(8) = &H8  '00
CBW(9) = &H0  '08
CBW(10) = &H0 '00
CBW(11) = &H0 '00

'///////////////  CBW Flag
CBW(12) = &H80                 '80

'////////////// LUN
CBW(13) = Lun                    '00

'///////////// CBD Len
CBW(14) = &HA                '0a

'////////////  UFI command

CBW(15) = &H25
CBW(16) = Lun * 32
 
CBW(17) = &H0         '00
CBW(18) = &H0        '00
CBW(19) = &H0        '00
CBW(20) = &H0         '40

'/////////////  Reverve
CBW(21) = 0

'//////////// Transfer Len

 
CBW(22) = &H0     '00
CBW(23) = &H0     '04

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
 Read_CapacityAU6371 = 0
 Exit Function
End If

'2. Readdata stage
 
result = ReadFile _
         (ReadHandle, _
          ReadData(0), _
         CBWDataTransferLength, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

If result = 0 Then
  Read_CapacityAU6371 = 0
  Exit Function  'Allen 0901
End If


For i = 0 To CBWDataTransferLength - 1
Debug.Print "k", i, Hex(ReadData(i)), Capacity(i)
'If ReadData(i) <> Capacity(i) Then
  
 ' Read_Capacity = 2  ' card format capacity has problem
  'Exit Function
'End If


Next i


'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
          
          
  
 
If result = 0 Then
 Read_CapacityAU6371 = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
     Read_CapacityAU6371 = 0
Else
      Read_CapacityAU6371 = 1
   
End If

 
End Function



Public Function ClosePipe() As Integer
On Error Resume Next
CloseHandle (ReadHandle)

CloseHandle (WriteHandle)


End Function
   
Public Function TestUnitReadyReadOverCurrent(Lun As Byte) As Byte
Dim CBW(0 To 30) As Byte
Dim CSW(0 To 12) As Byte
Dim i As Integer
Dim NumberOfBytesWritten As Long
Dim NumberOfBytesRead As Long
Dim result As Long

     For i = 0 To 30
    
        CBW(i) = 0
    
    Next i

CBW(0) = &H55 'signature
CBW(1) = &H53
CBW(2) = &H42
CBW(3) = &H43


CBW(4) = &H1  'package ID
CBW(5) = &H2
CBW(6) = &H3
CBW(7) = &H4


CBW(8) = &H0
CBW(9) = &H0
CBW(10) = &H0
CBW(11) = &H0


CBW(12) = &H80 '    CBW FLAG 0000
'CBW(13) = &H0
CBW(13) = Lun
CBW(14) = &H0
CBW(15) = &H0
CBW(16) = Lun * 32


'1. CBW output

TestUnitReadyReadOverCurrent = 7
 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)
 
 If result = 0 Then
     TestUnitReadyReadOverCurrent = 7
    Exit Function
 End If
 Call MsecDelay(0.15)
 CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
 
 Tester.Print "overCurrent value="; Hex(LightOff); " : 80:pass;00< 3.17:C0>3.43"

  
 
  If LightOff = 128 Then  ' clamp between 3.45 and 3.15
  
      Call PowerSet2(1, "2.5", "0.5", 1, "2.5", "0.5", 1)
       Call MsecDelay(0.8)
      CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
      Tester.Print "check low resistor ="; Hex(LightOff); " C0 :pass"
      If LightOff = &HC0 Then
          
          Call PowerSet2(1, "3.5", "0.5", 1, "3.5", "0.5", 1)
          Call MsecDelay(0.8)
          CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
           Tester.Print "check high resistor ="; Hex(LightOff); " 00 :pass"
           If LightOff = 0 Then
           TestUnitReadyReadOverCurrent = 5
            End If
       End If
  
  
     
    Exit Function
 End If
 
  If LightOff = 0 Then
     TestUnitReadyReadOverCurrent = 6
    Exit Function
 End If
 

 

 
 
 
'Call MsecDelay(0.2)
'2. CSW input
 'result = ReadFile _
 '        (ReadHandle, _
 '         CSW(0), _
 '         13, _
 '         NumberOfBytesRead, _
 '         HIDOverlapped)  'in
 
 
Exit Function

If result = 0 Then
Tester.Print "5"
  '  TestUnitReadyReadOverCurrent = 0
  '  result = ReadFile _
  '       (ReadHandle, _
  '        CSW(0), _
  '        13, _
  '        NumberOfBytesRead, _
  '        HIDOverlapped)  'in
          
  '   result = ReadFile _
  '       (ReadHandle, _
  '        CSW(0), _
  '        13, _
  '        NumberOfBytesRead, _
  '        HIDOverlapped)  'in
          
    Exit Function
End If

 
'3 CSW Status
If CSW(12) = 1 Then
    TestUnitReadyReadOverCurrent = 0
    
    Else
    TestUnitReadyReadOverCurrent = 1
End If

End Function

Public Sub MultiSlotTestAU6376()
   
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
 If ChipName = "AU6370DLF20" Or ChipName = "AU6378ALF20" Then
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
                  
 
                  
 If TmpChip = "AU6378ALF20" Then
                    
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
                If TmpChip = "AU6377ALF25" Then
                  VidName = "vid_1984"
                Else
                 VidName = "vid_058f"
                End If
                
                
                ' for FITIPOwer the VID name
                
                 VidName = "vid_18e3"
                
              
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
                
                
                If TmpChip = "AU6378ALF20" Then
                
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
                   
                    If TmpChip = "AU6378ALF20" Then
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
                 
                 If rv0 = 1 And Left(TmpChip, 10) = "AU6375HLF2" Then
                 
                    ClosePipe
                    rv0 = CBWTest_New_21_Sector_AU6377(0, 1)
                 
                   
                    ClosePipe
                    
                    
                     TmpLBA = LBA
                     LBA = 99
                         For i = 1 To 30
                             rv1 = 0
                             LBA = LBA + 199
                            
                             ClosePipe
                             rv1 = CBWTest_New_128_Sector_AU6375(0, 1)  ' write
                             ClosePipe
                             If rv1 <> 1 Then
                              LBA = TmpLBA
                             GoTo AU6377ALFResult
                             End If
                         Next
                      LBA = TmpLBA
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
               
                If rv2 = 1 And TmpChip = "AU6378ALF20" Then         '--- for SMC
                
                CardResult = DO_WritePort(card, Channel_P1A, &H18)  ' 0110 0100
                Call MsecDelay(0.5)
                ClosePipe
                rv2 = CBWTest_New(2, rv2, VidName)
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
                 End If
                 
              If rv2 = 1 And Left(TmpChip, 10) = "AU6377ALF2" And TmpChip <> "AU6377ALF21" Then           '--- for SMC
                
                CardResult = DO_WritePort(card, Channel_P1A, &H8 + RomSelector) ' 0110 0100
                Call MsecDelay(0.5)
                ClosePipe
                rv2 = CBWTest_New(2, rv2, VidName)
                Call LabelMenu(2, rv2, rv1)
                ClosePipe
              End If
                
                
               If rv2 = 1 And (TmpChip = "AU6376JLF20") Then      '--- for SMC
                
                  CardResult = DO_WritePort(card, Channel_P1A, &H18)   ' 0110 0100
                  Call MsecDelay(0.5)
                  CardResult = DO_WritePort(card, Channel_P1A, &H10)   ' 0110 0100
                  Call MsecDelay(0.5)
                 ClosePipe
                 rv2 = CBWTest_New(2, rv2, VidName)
                 Call LabelMenu(2, rv2, rv1)
                 ClosePipe
               End If
                
               '=============== SMC test END ==================================================
               
               rv3 = CBWTest_New(3, rv2, VidName)
               ClosePipe
               Call LabelMenu(3, rv3, rv2)
           
                 
               If Left(TmpChip, 10) = "AU6375HLF2" Then
               
                 If rv0 = 1 Then
                   
                    ClosePipe
                     rv0 = CBWTest_New_AU6375IncPattern(0, 1, VidName)
                     Call LabelMenu(0, rv0, 1)
                     ClosePipe
                 End If
                
                End If
                
                 
                 
                If Left(TmpChip, 10) = "AU6377ALF2" Then
                
                ' test chip
                         ClosePipe
                         rv4 = CBWTest_New(4, rv3, VidName)   'MMC test
                          Call LabelMenu(10, rv4, rv3)
                          ClosePipe
          
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If LightOff <> 127 Then
                            UsbSpeedTestResult = GPO_FAIL
                            rv4 = 2
                         End If
          
                End If
                 
                If TmpChip = "AU6378ALF20" Then
                
             
                         ClosePipe
                         rv4 = CBWTest_New(4, rv3, VidName)
                          Call LabelMenu(10, rv4, rv3)
                          ClosePipe
          
                         CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                         If rv4 = 1 And LightOff <> 252 Then
                            UsbSpeedTestResult = GPO_FAIL
                            rv4 = 2
                         End If
          
                End If
                 
                 
                    
                  If ChipName = "AU6376" And TmpChip = "AU6370DLF20" Then
                  Call MsecDelay(0.1)
                  rv4 = 1
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 If LightOff <> 254 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                 End If
                     Call LabelMenu(3, rv4, rv3)
                     
                        
                 End If
                 
                 
                 If ChipName = "AU6368A1" Then
                 Call MsecDelay(0.1)
                  rv4 = 1
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 If LightOff <> 192 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                 End If
                     Call LabelMenu(3, rv4, rv3)
                     
                       CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 
                 End If
                 
                  If ChipName = "AU6376" And (Left(TmpChip, 10) <> "AU6377ALF2" And TmpChip <> "AU6378ALF20") Then
                  Call MsecDelay(0.1)
                  rv4 = 1
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                               If TmpChip = "AU6370DLF20" Or TmpChip = "AU6376JLF20" Then
                               
                                  If LightOff <> 254 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                 End If
                               Else
                                 If LightOff <> 252 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                  rv4 = 2
                                 End If
                              End If
                     Call LabelMenu(3, rv4, rv3)
                     
                        
                 End If
                 
                If ChipName = "AU6368A" Then
                    If rv3 = 1 Then
                           CardResult = DO_WritePort(card, Channel_P1A, &H74)  ' 0111 0100
                           Call MsecDelay(0.1)
                           CardResult = DO_WritePort(card, Channel_P1A, &H54)  ' 0101 0100
                           Call MsecDelay(0.1)
                           rv4 = CBWTest_New(3, rv3, VidName)
                             ClosePipe
                           If rv4 = 1 Then
                                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                 If LightOff <> 132 Then
                                   UsbSpeedTestResult = GPO_FAIL
                                    rv4 = 2
                                 End If
                             End If
                         Else
                         rv4 = 4
                         End If
                         Call LabelMenu(3, rv4, rv3)
                 End If
                
                
                Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
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
                        ElseIf rv2 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                        ElseIf rv2 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                         ElseIf rv3 = WRITE_FAIL Or rv4 = WRITE_FAIL Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                        ElseIf rv3 = READ_FAIL Or rv4 = READ_FAIL Then
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
Public Function CBWTest_New_AU9254(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

   CBWDataTransferLength = 1024
 
'   For i = 0 To CBWDataTransferLength - 1
    
'         ReadData(i) = 0

'   Next

    If PreSlotStatus <> 1 Then
        CBWTest_New_AU9254 = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New_AU9254 = 0
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
      CBWTest_New_AU9254 = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    
    
  '  If Vid_PID = "6335" Or Vid_PID = "9369" Or Vid_PID = "6331" Or Vid_PID = "9462" Then
  '     CBWTest_New_AU9254 = 1   ' usb 2.0 high speed fail
  '      Exit Function
  ' End If
    
    
       If Vid_PID = "9462" Or Vid_PID = "413c" Then
        CBWTest_New_AU9254 = 1   ' usb 2.0 high speed fail
         Exit Function
   End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_New_AU9254 = 2   ' Write fail
      Exit Function
    End If
 
    '======================================
    
    If Vid_PID = "6377" And HubPort = 0 Then
      TmpInteger = TestUnitSpeed(Lun)
    
      If TmpInteger = 0 Then
        
         CBWTest_New_AU9254 = 2   ' usb 2.0 high speed fail
         UsbSpeedTestResult = 2
          Exit Function
      End If
       CBWTest_New_AU9254 = 1
      Exit Function
    End If

    
     ' for unitSpeed
  If Vid_PID = "6335" Then
      TmpInteger = TestUnitSpeed(Lun)
    
      If TmpInteger = 0 Then
        
         CBWTest_New_AU9254 = 2   ' usb 2.0 high speed fail
         UsbSpeedTestResult = 2
          Exit Function
      End If
       CBWTest_New_AU9254 = 1
      Exit Function
    End If
    
   TmpInteger = TestUnitReady(Lun)
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
           CBWTest_New_AU9254 = 2  'Write fail
          Exit Function
        End If
        
    End If

    '======================================
    If ChipName = "AU6371" Or Vid_PID = "0054" Then
        TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
    End If
     TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
   ' TmpInteger = Read_Data1(Lba, Lun, CBWDataTransferLength)
    TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
      
    If TmpInteger = 0 Then
         CBWTest_New_AU9254 = 2  'write fail
           Exit Function
     End If
    
      
    TmpInteger = Write_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_AU9254 = 2  'write fail
        Exit Function
    End If
    
    TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_AU9254 = 3    'Read fail
        Exit Function
    End If
     
    For i = 0 To CBWDataTransferLength - 1
    
        If ReadData(i) <> Pattern(i) Then
          CBWTest_New_AU9254 = 3    'Read fail
          Exit Function
        End If
    
    Next
    
  CBWTest_New_AU9254 = 1
        
    
    End Function


Public Function CBWTest_New_AU6350CF(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
Dim i As Integer
Dim WriteTest As Integer
Dim TmpString As String
Dim TmpInteger As Integer
Dim TimerCounter As Integer
Dim OldTimer
Dim CBWDataTransferLength As Long

   CBWDataTransferLength = 1024
 
'   For i = 0 To CBWDataTransferLength - 1
    
'         ReadData(i) = 0

'   Next

    If PreSlotStatus <> 1 Then
        CBWTest_New_AU6350CF = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New_AU6350CF = 0
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
      CBWTest_New_AU6350CF = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    
    
  '  If Vid_PID = "6335" Or Vid_PID = "9369" Or Vid_PID = "6331" Or Vid_PID = "9462" Then
  '     CBWTest_New_AU6350CF = 1   ' usb 2.0 high speed fail
  '      Exit Function
  ' End If
    
    
   
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_New_AU6350CF = 2   ' Write fail
      Exit Function
    End If
 
    '======================================
    
   
      TmpInteger = TestUnitSpeed(Lun)
    
      If TmpInteger = 0 Then
        
         CBWTest_New_AU6350CF = 2   ' usb 2.0 high speed fail
         UsbSpeedTestResult = 2
          Exit Function
      End If
     

    
     ' for unitSpeed
  
   TmpInteger = TestUnitReady(Lun)
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
           CBWTest_New_AU6350CF = 2  'Write fail
          Exit Function
        End If
        
    End If

    '======================================
    
        TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
    
     TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
   ' TmpInteger = Read_Data1(Lba, Lun, CBWDataTransferLength)
      TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
      
    If TmpInteger = 0 Then
         CBWTest_New_AU6350CF = 2  'write fail
           Exit Function
     End If
    
      
    TmpInteger = Write_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_AU6350CF = 2  'write fail
        Exit Function
    End If
    
    TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_AU6350CF = 3    'Read fail
        Exit Function
    End If
     
    For i = 0 To CBWDataTransferLength - 1
    
        If ReadData(i) <> Pattern(i) Then
          CBWTest_New_AU6350CF = 3    'Read fail
          Exit Function
        End If
    
    Next
    
  CBWTest_New_AU6350CF = 1
        
    
    End Function
Public Function AU6435Close_OverCurrent(PreviousResult As Byte) As Byte

If PreviousResult <> 1 Then
AU6435Close_OverCurrent = 4
Exit Function
End If

Dim CBW(0 To 30) As Byte
Dim CSW(0 To 12) As Byte
Dim i As Integer
Dim NumberOfBytesWritten As Long
Dim NumberOfBytesRead As Long
Dim result As Long

     For i = 0 To 30
    
        CBW(i) = &H0
    
    Next i

CBW(0) = &H55 'signature
CBW(1) = &H53
CBW(2) = &H42
CBW(3) = &H43


CBW(4) = &H1  'package ID
CBW(5) = &H2
CBW(6) = &H3
CBW(7) = &H4


CBW(8) = &H0
CBW(9) = &H0
CBW(10) = &H0
CBW(11) = &H0


CBW(12) = &H80 '    CBW FLAG 0000
CBW(13) = 0 'Lun
CBW(14) = &H10
CBW(15) = &HC7
CBW(16) = &H1F
CBW(17) = &H5
CBW(18) = &H8F
CBW(19) = &HC7
CBW(20) = &H1
CBW(21) = &H0

 
'Dim result As Long
'1. CBW output

AU6435Close_OverCurrent = 0

result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)


If result = 0 Then
    AU6435Close_OverCurrent = 0
    Exit Function
End If


'2. CSW input
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

'If result = 0 Then
'   AU6435Close_OverCurrent = 1
'   Exit Function
'End If

 
'3 CSW Status
If CSW(12) = 0 Then
    'For i = 0 To 12
    '    Debug.Print CSW(i)
    'Next
    AU6435Close_OverCurrent = 1

    Else
    AU6435Close_OverCurrent = 0
End If

End Function

Public Function AU6435Set_Pad_Driving27(PreviousResult As Byte) As Byte

If PreviousResult <> 1 Then
AU6435Set_Pad_Driving27 = 4
Exit Function
End If

Dim CBW(0 To 30) As Byte
Dim CSW(0 To 12) As Byte
Dim i As Integer
Dim NumberOfBytesWritten As Long
Dim NumberOfBytesRead As Long
Dim result As Long

     For i = 0 To 30
    
        CBW(i) = &H0
    
    Next i

CBW(0) = &H55 'signature
CBW(1) = &H53
CBW(2) = &H42
CBW(3) = &H43


CBW(4) = &H1  'package ID
CBW(5) = &H2
CBW(6) = &H3
CBW(7) = &H4


CBW(8) = &H0
CBW(9) = &H0
CBW(10) = &H0
CBW(11) = &H0


CBW(12) = &H80 '    CBW FLAG 0000
CBW(13) = 0 'Lun
CBW(14) = &H10
CBW(15) = &HC7
CBW(16) = &H1F
CBW(17) = &H5
CBW(18) = &H8F
CBW(19) = &HC7
CBW(20) = &H85
CBW(21) = &H53
CBW(22) = &H44
CBW(23) = &H42  'Pad Driving 4.5mA/3.6mA/2.7mA  0x22/0x32/0x42
 
'Dim result As Long
'1. CBW output

AU6435Set_Pad_Driving27 = 0

result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)


If result = 0 Then
    AU6435Set_Pad_Driving27 = 0
    Exit Function
End If


'2. CSW input
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

'If result = 0 Then
'   AU6435Close_OverCurrent = 1
'   Exit Function
'End If

 
'3 CSW Status
If CSW(12) = 0 Then
    For i = 0 To 12
        Debug.Print CSW(i)
    Next
    AU6435Set_Pad_Driving27 = 1

    Else
    AU6435Set_Pad_Driving27 = 0
End If

End Function

Public Function AU6435Set_Pad_Driving54(PreviousResult As Byte) As Byte

If PreviousResult <> 1 Then
AU6435Set_Pad_Driving54 = 4
Exit Function
End If

Dim CBW(0 To 30) As Byte
Dim CSW(0 To 12) As Byte
Dim i As Integer
Dim NumberOfBytesWritten As Long
Dim NumberOfBytesRead As Long
Dim result As Long
Dim CBWDataTransferLen(0 To 3) As Byte
Dim OutData(0 To 511) As Byte
Dim CBWDataTransferLength As Long
Dim tmpV(0 To 2) As Byte

CBWDataTransferLength = 512

For i = 0 To 30
    CBW(i) = &H0
Next i

For i = 1 To CBWDataTransferLength - 1
    OutData(i) = 0
Next

'result = TestUnitReady(Lun)

'If result = 0 Then
'    result = RequestSense(Lun)
'End If

CBW(0) = &H55 'signature
CBW(1) = &H53
CBW(2) = &H42
CBW(3) = &H43


CBW(4) = &H1  'package ID
CBW(5) = &H2
CBW(6) = &H3
CBW(7) = &H4


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

CBW(12) = &H0  '    CBW FLAG 0000
CBW(13) = 0 'Lun
CBW(14) = &H10
CBW(15) = &HC7
CBW(16) = &H1F
CBW(17) = &H5
CBW(18) = &H8F
CBW(19) = &HC7
CBW(20) = &H8
CBW(21) = &H30
CBW(22) = &H35
CBW(23) = &H38
CBW(24) = &H46

OutData(0) = &H3
OutData(1) = &H5
OutData(2) = &H3
OutData(3) = &H52
OutData(4) = &H3
OutData(5) = &H7
OutData(6) = &H3
OutData(7) = &H28
OutData(8) = &H0



'Dim result As Long
'1. CBW output
 
result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)    'out

If result = 0 Then
    AU6435Set_Pad_Driving54 = 0
    Exit Function
End If


'==================================
result = WriteFile _
       (WriteHandle, _
       OutData(0), _
       CBWDataTransferLength, _
       NumberOfBytesWritten, _
       0)    'out
'==================================
'result = RequestSense(Lun)
If result = 0 Then
    AU6435Set_Pad_Driving54 = 0
'    Exit Function
End If



'3 . CSW
result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in


If result = 0 Then
    AU6435Set_Pad_Driving54 = 0
    Exit Function
End If
 
'For i = 1 To 511
'    Debug.Print i; ReadData(i)
'Next
 
 
 
If CSW(12) = 1 Then
    AU6435Set_Pad_Driving54 = 0
Else
    AU6435Set_Pad_Driving54 = 1
End If

End Function

Public Function AU6435ForceSDHC(PreviousResult As Byte) As Byte

If PreviousResult <> 1 Then
    AU6435ForceSDHC = 4
    Exit Function
End If

Dim CBW(0 To 30) As Byte
Dim CSW(0 To 12) As Byte
Dim i As Integer
Dim NumberOfBytesWritten As Long
Dim NumberOfBytesRead As Long
Dim result As Long

     For i = 0 To 30
    
        CBW(i) = &H0
    
    Next i

CBW(0) = &H55 'signature
CBW(1) = &H53
CBW(2) = &H42
CBW(3) = &H43


CBW(4) = &H1  'package ID
CBW(5) = &H2
CBW(6) = &H3
CBW(7) = &H4


CBW(8) = &H0
CBW(9) = &H0
CBW(10) = &H0
CBW(11) = &H0


CBW(12) = &H80 '    CBW FLAG 0000
CBW(13) = 0 'Lun
CBW(14) = &H10
CBW(15) = &HC7
CBW(16) = &H1F
CBW(17) = &H5
CBW(18) = &H8F
CBW(19) = &HC7
CBW(20) = &H84
CBW(21) = &H1
CBW(22) = &H0

'Dim result As Long
'1. CBW output

AU6435ForceSDHC = 0

result = WriteFile _
       (WriteHandle, _
       CBW(0), _
       31, _
       NumberOfBytesWritten, _
       0)


If result = 0 Then
    AU6435ForceSDHC = 0
    Exit Function
End If


'2. CSW input
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in

 
'3 CSW Status
If CSW(12) = 0 Then
    For i = 0 To 12
        Debug.Print CSW(i)
    Next
    AU6435ForceSDHC = 1

    Else
    AU6435ForceSDHC = 0
End If

End Function


