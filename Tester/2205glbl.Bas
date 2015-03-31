Attribute VB_Name = "AU3130"
Option Explicit

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long

Global Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
'Global card As Integer
Global card_number As Integer
Global card_type As Integer
Global ch_num As Integer
Global ch_range As Integer
Global ch_ref As Integer

Dim ch_cnt As Integer
Global gnBuffer(4000) As Integer
Global gnBufferWMA(4000) As Integer
Dim count1 As Long
Dim dma_size As Long
Global Card2214 As Integer
Global RecordMode As Integer
Global RecordModeWMA As Integer
Global RecordModeMP3 As Integer
Global Card2214Exist As Integer
Public TmpChip As String
Public OldLBa As Long
'============================== for Audio mode =================
Global ADCch_num(0 To 1) As Integer
Global ADCch_range(0 To 1) As Integer
Global ADCch_ref(0 To 1) As Integer

Global ADCBuffer(0 To 11999) As Integer
 
Global MP3MatchIndex As Byte
Global WMAMatchIndex As Byte
Global MP3PatternIndex As Byte
Global WMAPatternIndex As Byte
Global PatternSlope(0 To 5999) As Long
Global RawDataSlope(0 To 5999) As Long

Global PatternAcc(0 To 5999) As Long
Global RawDataAcc(0 To 5999) As Long
Global SineCurve(0 To 1, 0 To 5999) As Long
Global SineCurve2(0 To 1, 0 To 5999) As Long


 
Global MP3RawData(0 To 1, 0 To 5999) As Integer
Global MP3Pattern(0 To 20, 0 To 1, 0 To 5999) As Integer

Global WMARawData(0 To 1, 0 To 5999) As Integer
Global WMAPattern(0 To 10, 0 To 1, 0 To 5999) As Integer

Global LoadAudioPatternFlag As Byte
Global MP3CompareResult As Byte
Global WMACompareResult As Byte

Global MP3DistinguishResult As Byte
Global WMADistinguishResult As Byte
 


Public Sub AU3152HLF2C_1_SlotTest_Noise()
Dim i As Integer
  
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                 '=========================================
                '    POWER on
                '=========================================
               '   CardResult = DO_WritePort(card, Channel_P1A, &H80)
                 '    CardResult = DO_WritePort(card, Channel_P1A, &HC0)  ' Allen debug
                   CardResult = DO_WritePort(card, Channel_P1A, &H80)  ' Allen debug
                 'If CardResult <> 0 Then
                 '   MsgBox "Power off fail"
                 '   End
                ' End If
                 
                  Call MsecDelay(0.05)
                CardResult = DO_WritePort(card, Channel_P1A, &H0)
              '    CardResult = DO_WritePort(card, Channel_P1A, &H2F)   ' usb Power Off
                  Call MsecDelay(0.8)    'power on time
                  
             '   Call MsecDelay(1#)      'power on time
                  
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
            
                ClosePipe
                LBA = LBA + 1
                OldLBa = LBA
       
                
                LBA = LBA + 40000 ' 20000 mp3 shift
                
                rv0 = CBWTest_New_31303(0, 1, "vid_058f")
                LBA = OldLBa
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 Tester.Cls
                Tester.Print "LBA="; LBA
                Tester.Print "rv0="; rv0
                 
                If rv0 = 1 Then
                    
                          '=================================
                          '   begin load pattern
                          '=================================
                          
                         
                          If OldChipName <> ChipName Then
                                 
                             '   Call LoadAU3150JLF27AudioPattern
                                OldChipName = ChipName
                          End If
                                                      
              
                          
                          '==========================================
                          '  do a low-high-low pulse at play button
                          '==========================================
                         
                             
                             
                          
                              CardResult = DO_WritePort(card, Channel_P1A, &H40)
                                 Call MsecDelay(2.5)
                            
                          
                          '==========================================
                          '  Get ADC Data and compare result
                          '==========================================
                   Call MsecDelay(2.2)
                           Call DACGetData2ChannelUpTriggleAU3152HL
                           Call ADCBuffer2AudioBuffer(0)
                        '   Call CompareDataTrace(0)
                           Call DistingDataTrace(0)
                           
                           rv2 = MP3DistinguishResult
                  Else
                   rv2 = 4
                 End If
                 
                 Call LabelMenu(2, rv2, rv0)
                 

               
                If (rv0 = 1) And (rv2 = 1 Or Tester.AudioRecordCheck.Value = 1) Then
                  
                      D2K_Release_Card (Card2214)
                     Card2214Exist = 0
                     
                    '=========================================
                    '
                    '   Start Next one song WMA
                    '
                    '========================================
                     Tester.Print "next one press"
                      CardResult = DO_WritePort(card, Channel_P1A, &H41)
                    
                      Call MsecDelay(0.2)
                      CardResult = DO_WritePort(card, Channel_P1A, &H40)
                  
                      Tester.Print "next one up"
                        
                        
                         
                      '=========================================
                      '
                      '   End Next one song WMA
                      '
                      '========================================
                    
                     ' Call MsecDelay(1#)
                     '  CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                       Call DACGetData2ChannelUpTriggleAU3152HL
                       Call ADCBuffer2AudioBuffer(1)
                      ' Call CompareDataTrace(1)
                        Call DistingDataTrace(1)
                           
                           rv3 = WMADistinguishResult
                 Else
                   rv3 = 4
                 End If
                 Call LabelMenu(3, rv3, rv2)
                  
                 '========================================================
                 '
                 '  Recode Raw data to pattern file
                 '
                 '========================================================
                  
                  
                 If Tester.AudioRecordCheck.Value = 1 Then
                                       
                     Open App.Path & "\MP3Pattern\MP3.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, MP3RawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, MP3RawData(1, i)
                         Next i
                      
    
                    Close #4
                
                
                                  
                     Open App.Path & "\WMAPattern\WMA.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, WMARawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, WMARawData(1, i)
                         Next i
                      
    
                    Close #4
                    
                 End If
                 
                 If rv2 * rv3 = 1 Then
                 
                   ' Tester.Cls
                    Tester.Print "MP3Matchmode="; MP3MatchIndex
                    Tester.Print "WMAMatchmode="; WMAMatchIndex
                 End If

                 dma2205_4.Show
                 dma2205_4.ShowData3 0, 6000
                
                 Card2214Exist = 0
                  D2K_Release_Card (Card2214)
                  CardResult = DO_WritePort(card, Channel_P1A, &HC0)  ' Allen debug
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
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
                ElseIf rv0 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub

Public Sub AU3152HLF2E_1_SlotTest_Noise()
Dim i As Integer
  
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                GoTo Test_MP3:
                
TEST_USB:
                 '=========================================
                '    POWER on
                '=========================================
               '   CardResult = DO_WritePort(card, Channel_P1A, &H80)
                      CardResult = DO_WritePort(card, Channel_P1A, &HC0)  ' Allen debug
                       Call MsecDelay(0.05)
                   CardResult = DO_WritePort(card, Channel_P1A, &H80)  ' Allen debug
                 'If CardResult <> 0 Then
                 '   MsgBox "Power off fail"
                 '   End
                ' End If
                 
                  Call MsecDelay(0.05)
                CardResult = DO_WritePort(card, Channel_P1A, &H0)
              '    CardResult = DO_WritePort(card, Channel_P1A, &H2F)   ' usb Power Off
                  Call MsecDelay(0.8)    'power on time
                  
             '   Call MsecDelay(1#)      'power on time
                  
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
            
                ClosePipe
                LBA = LBA + 1
                OldLBa = LBA
       
                
                LBA = LBA + 40000 ' 20000 mp3 shift
                
                rv0 = CBWTest_New_31303(0, 1, "vid_058f")
                LBA = OldLBa
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 Tester.Cls
                Tester.Print "LBA="; LBA
                Tester.Print "rv0="; rv0
                
                
                
                GoTo TEST_RESULT:
                 
Test_MP3:
                rv0 = 1
                If rv0 = 1 Then
                    
                          '=================================
                          '   begin load pattern
                          '=================================
                          
                         
                          If OldChipName <> ChipName Then
                                 
                             '   Call LoadAU3150JLF27AudioPattern
                                OldChipName = ChipName
                          End If
                                                      
              
                          
                          '==========================================
                          '  do a low-high-low pulse at play button
                          '==========================================
                         
                               CardResult = DO_WritePort(card, Channel_P1A, &HC0)
                                 Call MsecDelay(0.05)
                             
                          
                              CardResult = DO_WritePort(card, Channel_P1A, &H40)
                                 Call MsecDelay(2.5)
                            
                          
                          '==========================================
                          '  Get ADC Data and compare result
                          '==========================================
                          Call MsecDelay(2.2)
                           Call DACGetData2ChannelUpTriggleAU3152HL
                           Call ADCBuffer2AudioBuffer(0)
                        '   Call CompareDataTrace(0)
                           Call DistingDataTrace(0)
                           
                           rv2 = MP3DistinguishResult
                  Else
                   rv2 = 4
                 End If
                 
                 Call LabelMenu(2, rv2, rv0)
                 

               
                If (rv0 = 1) And (rv2 = 1 Or Tester.AudioRecordCheck.Value = 1) Then
                  
                      D2K_Release_Card (Card2214)
                     Card2214Exist = 0
                     
                    '=========================================
                    '
                    '   Start Next one song WMA
                    '
                    '========================================
                     Tester.Print "next one press"
                      CardResult = DO_WritePort(card, Channel_P1A, &H41)
                    
                      Call MsecDelay(0.2)
                      CardResult = DO_WritePort(card, Channel_P1A, &H40)
                  
                      Tester.Print "next one up"
                        
                        
                         
                      '=========================================
                      '
                      '   End Next one song WMA
                      '
                      '========================================
                    
                     ' Call MsecDelay(1#)
                     '  CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                       Call DACGetData2ChannelUpTriggleAU3152HL
                       Call ADCBuffer2AudioBuffer(1)
                      ' Call CompareDataTrace(1)
                        Call DistingDataTrace(1)
                           
                           rv3 = WMADistinguishResult
                 Else
                   rv3 = 4
                 End If
                 Call LabelMenu(3, rv3, rv2)
                  
                 '========================================================
                 '
                 '  Recode Raw data to pattern file
                 '
                 '========================================================
                  
                  
                 If Tester.AudioRecordCheck.Value = 1 Then
                                       
                     Open App.Path & "\MP3Pattern\MP3.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, MP3RawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, MP3RawData(1, i)
                         Next i
                      
    
                    Close #4
                
                
                                  
                     Open App.Path & "\WMAPattern\WMA.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, WMARawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, WMARawData(1, i)
                         Next i
                      
    
                    Close #4
                    
                 End If
                 
                 If rv2 * rv3 = 1 Then
                 
                   ' Tester.Cls
                    Tester.Print "MP3Matchmode="; MP3MatchIndex
                    Tester.Print "WMAMatchmode="; WMAMatchIndex
                 End If

                 dma2205_4.Show
                 dma2205_4.ShowData3 0, 6000
                
                 Card2214Exist = 0
                  D2K_Release_Card (Card2214)
                  CardResult = DO_WritePort(card, Channel_P1A, &HC0)  ' Allen debug
                  
                  
                   rv0 = 0
                  GoTo TEST_USB:
                   
TEST_RESULT:
                  
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
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
                             
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv1 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                 ElseIf rv1 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
                    TestResult = "CF_RF"
             
                ElseIf rv0 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
                
                  CardResult = DO_WritePort(card, Channel_P1A, &HC0)
                                  
End Sub
Public Sub AU3152HLF2D_1_SlotTest_Noise()
Dim i As Integer
  
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                 '=========================================
                '    POWER on
                '=========================================
               '   CardResult = DO_WritePort(card, Channel_P1A, &H80)
                 '    CardResult = DO_WritePort(card, Channel_P1A, &HC0)  ' Allen debug
                   CardResult = DO_WritePort(card, Channel_P1A, &H80)  ' Allen debug
                 'If CardResult <> 0 Then
                 '   MsgBox "Power off fail"
                 '   End
                ' End If
                 
                  Call MsecDelay(0.05)
                CardResult = DO_WritePort(card, Channel_P1A, &H0)
              '    CardResult = DO_WritePort(card, Channel_P1A, &H2F)   ' usb Power Off
                  Call MsecDelay(0.8)    'power on time
                  
             '   Call MsecDelay(1#)      'power on time
                  
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
            
                ClosePipe
                LBA = LBA + 1
                OldLBa = LBA
       
                
                LBA = LBA + 40000 ' 20000 mp3 shift
                
                rv0 = CBWTest_New_31303(0, 1, "vid_058f")
                LBA = OldLBa
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 Tester.Cls
                Tester.Print "LBA="; LBA
                Tester.Print "rv0="; rv0
                 
                If rv0 = 1 Then
                    
                          '=================================
                          '   begin load pattern
                          '=================================
                          
                         
                          If OldChipName <> ChipName Then
                                 
                             '   Call LoadAU3150JLF27AudioPattern
                                OldChipName = ChipName
                          End If
                                                      
              
                          
                          '==========================================
                          '  do a low-high-low pulse at play button
                          '==========================================
                         
                             
                             
                          
                              CardResult = DO_WritePort(card, Channel_P1A, &H40)
                                 Call MsecDelay(2.5)
                            
                          
                          '==========================================
                          '  Get ADC Data and compare result
                          '==========================================
                   Call MsecDelay(2.2)
                           Call DACGetData2ChannelUpTriggleAU3152HL
                           Call ADCBuffer2AudioBuffer(0)
                        '   Call CompareDataTrace(0)
                           Call DistingDataTrace(0)
                           
                           rv2 = MP3DistinguishResult
                  Else
                   rv2 = 4
                 End If
                 
                 Call LabelMenu(2, rv2, rv0)
                 

               
                If (rv0 = 1) And (rv2 = 1 Or Tester.AudioRecordCheck.Value = 1) Then
                  
                      D2K_Release_Card (Card2214)
                     Card2214Exist = 0
                     
                    '=========================================
                    '
                    '   Start Next one song WMA
                    '
                    '========================================
                     Tester.Print "next one press"
                      CardResult = DO_WritePort(card, Channel_P1A, &H41)
                    
                      Call MsecDelay(0.2)
                      CardResult = DO_WritePort(card, Channel_P1A, &H40)
                  
                      Tester.Print "next one up"
                        
                        
                         
                      '=========================================
                      '
                      '   End Next one song WMA
                      '
                      '========================================
                    
                     ' Call MsecDelay(1#)
                     '  CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                       Call DACGetData2ChannelUpTriggleAU3152HL
                       Call ADCBuffer2AudioBuffer(1)
                      ' Call CompareDataTrace(1)
                        Call DistingDataTrace(1)
                           
                           rv3 = WMADistinguishResult
                 Else
                   rv3 = 4
                 End If
                 Call LabelMenu(3, rv3, rv2)
                  
                 '========================================================
                 '
                 '  Recode Raw data to pattern file
                 '
                 '========================================================
                  
                  
                 If Tester.AudioRecordCheck.Value = 1 Then
                                       
                     Open App.Path & "\MP3Pattern\MP3.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, MP3RawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, MP3RawData(1, i)
                         Next i
                      
    
                    Close #4
                
                
                                  
                     Open App.Path & "\WMAPattern\WMA.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, WMARawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, WMARawData(1, i)
                         Next i
                      
    
                    Close #4
                    
                 End If
                 
                 If rv2 * rv3 = 1 Then
                 
                   ' Tester.Cls
                    Tester.Print "MP3Matchmode="; MP3MatchIndex
                    Tester.Print "WMAMatchmode="; WMAMatchIndex
                 End If

                 dma2205_4.Show
                 dma2205_4.ShowData3 0, 6000
                
                 Card2214Exist = 0
                  D2K_Release_Card (Card2214)
                  CardResult = DO_WritePort(card, Channel_P1A, &HC0)  ' Allen debug
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
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
                ElseIf rv0 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub

Public Sub AU3150QLF28_1_SlotTest_Noise()
Dim i As Integer
  
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                 '=========================================
                '    POWER on
                '=========================================
               '   CardResult = DO_WritePort(card, Channel_P1A, &H80)
                    CardResult = DO_WritePort(card, Channel_P1A, &HC0)  ' Allen debug
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                CardResult = DO_WritePort(card, Channel_P1A, &H6F)
              '    CardResult = DO_WritePort(card, Channel_P1A, &H2F)   ' usb Power Off
                  Call MsecDelay(0.8)    'power on time
                  
              '   Call MsecDelay(1#)      'power on time
                  
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
            
                ClosePipe
                LBA = LBA + 1
                OldLBa = LBA
       
                
                LBA = LBA + 40000 ' 20000 mp3 shift
                
                rv0 = CBWTest_New_31303(0, 1, "vid_058f")
                LBA = OldLBa
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 Tester.Cls
                Tester.Print "LBA="; LBA
                Tester.Print "rv0="; rv0
                 
                If rv0 = 1 Then
                    
                          '=================================
                          '   begin load pattern
                          '=================================
                          
                         
                          If OldChipName <> ChipName Then
                                 
                                Call LoadAU3150JLF27AudioPattern
                                OldChipName = ChipName
                          End If
                                                      
              
                          
                          '==========================================
                          '  do a low-high-low pulse at play button
                          '==========================================
                         
                             
                             
                             CardResult = DO_WritePort(card, Channel_P1A, &HCF)
                             
                             Call MsecDelay(0.3)
                             CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                                 Call MsecDelay(2.5)
                                'Call MsecDelay(4.5)
                              CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                            '   CardResult = DO_WritePort(card, Channel_P1A, &HDF)
                          '    Call MsecDelay(4)    'power on time and DAC initial time
                               
                            ' Call MsecDelay(2.5)     'for 3150K power on time
                               Call MsecDelay(0.3)
                             CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                            
                            ' CardResult = DO_WritePort(card, Channel_P1A, &HCF)
                          '==========================================
                          '  Get ADC Data and compare result
                          '==========================================
                   Call MsecDelay(1.5)
                           Call DACGetData2ChannelUpTriggle
                           Call ADCBuffer2AudioBuffer(0)
                           Call CompareDataTrace(0)
                           
                           rv2 = MP3CompareResult
                  Else
                   rv2 = 4
                 End If
                 
                 Call LabelMenu(2, rv2, rv0)
                 

               
                If (rv0 = 1) And (rv2 = 1 Or Tester.AudioRecordCheck.Value = 1) Then
                  
                      D2K_Release_Card (Card2214)
                     Card2214Exist = 0
                     
                    '=========================================
                    '
                    '   Start Next one song WMA
                    '
                    '========================================
                     Tester.Print "next one press"
                      CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                    
                      Call MsecDelay(0.2)
                      CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                  
                      Tester.Print "next one up"
                        
                        
                         
                      '=========================================
                      '
                      '   End Next one song WMA
                      '
                      '========================================
                    
                     ' Call MsecDelay(1#)
                     '  CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                       Call DACGetData2ChannelUpTriggle
                       Call ADCBuffer2AudioBuffer(1)
                       Call CompareDataTrace(1)
                           
                       rv3 = WMACompareResult
                 Else
                   rv3 = 4
                 End If
                 Call LabelMenu(3, rv3, rv2)
                  
                 '========================================================
                 '
                 '  Recode Raw data to pattern file
                 '
                 '========================================================
                  
                  
                 If Tester.AudioRecordCheck.Value = 1 Then
                                       
                     Open App.Path & "\MP3Pattern\MP3.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, MP3RawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, MP3RawData(1, i)
                         Next i
                      
    
                    Close #4
                
                
                                  
                     Open App.Path & "\WMAPattern\WMA.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, WMARawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, WMARawData(1, i)
                         Next i
                      
    
                    Close #4
                    
                 End If
                 
                 If rv2 * rv3 = 1 Then
                 
                   ' Tester.Cls
                    Tester.Print "MP3Matchmode="; MP3MatchIndex
                    Tester.Print "WMAMatchmode="; WMAMatchIndex
                 End If

                 dma2205_4.Show
                 dma2205_4.ShowData 0, 6000
                
                 Card2214Exist = 0
                  D2K_Release_Card (Card2214)
                  CardResult = DO_WritePort(card, Channel_P1A, &HC0)  ' Allen debug
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
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
                ElseIf rv0 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub
Public Sub AU3150KLF28_1_SlotTest_Noise()
Dim i As Integer
  
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                 '=========================================
                '    POWER on
                '=========================================
               '   CardResult = DO_WritePort(card, Channel_P1A, &H80)
                    CardResult = DO_WritePort(card, Channel_P1A, &HC0)  ' Allen debug
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                CardResult = DO_WritePort(card, Channel_P1A, &H6F)
              '    CardResult = DO_WritePort(card, Channel_P1A, &H2F)   ' usb Power Off
                  Call MsecDelay(0.8)    'power on time
                  
              '   Call MsecDelay(1#)      'power on time
                  
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
            
                ClosePipe
                LBA = LBA + 1
                OldLBa = LBA
       
                
                LBA = LBA + 40000 ' 20000 mp3 shift
                
                rv0 = CBWTest_New_31303(0, 1, "vid_058f")
                LBA = OldLBa
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 Tester.Cls
                Tester.Print "LBA="; LBA
                Tester.Print "rv0="; rv0
                 
                If rv0 = 1 Then
                    
                          '=================================
                          '   begin load pattern
                          '=================================
                          
                         
                          If OldChipName <> ChipName Then
                                 
                                Call LoadAU3150JLF27AudioPattern
                                OldChipName = ChipName
                          End If
                                                      
              
                          
                          '==========================================
                          '  do a low-high-low pulse at play button
                          '==========================================
                         
                             
                             
                             CardResult = DO_WritePort(card, Channel_P1A, &HCF)
                             
                             Call MsecDelay(0.3)
                             CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                                 Call MsecDelay(2.5)
                                'Call MsecDelay(4.5)
                              CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                            '   CardResult = DO_WritePort(card, Channel_P1A, &HDF)
                          '    Call MsecDelay(4)    'power on time and DAC initial time
                               
                            ' Call MsecDelay(2.5)     'for 3150K power on time
                               Call MsecDelay(0.3)
                             CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                            
                            ' CardResult = DO_WritePort(card, Channel_P1A, &HCF)
                          '==========================================
                          '  Get ADC Data and compare result
                          '==========================================
                   Call MsecDelay(1.5)
                           Call DACGetData2ChannelUpTriggle
                           Call ADCBuffer2AudioBuffer(0)
                           Call CompareDataTrace(0)
                           
                           rv2 = MP3CompareResult
                  Else
                   rv2 = 4
                 End If
                 
                 Call LabelMenu(2, rv2, rv0)
                 

               
                If (rv0 = 1) And (rv2 = 1 Or Tester.AudioRecordCheck.Value = 1) Then
                  
                      D2K_Release_Card (Card2214)
                     Card2214Exist = 0
                     
                    '=========================================
                    '
                    '   Start Next one song WMA
                    '
                    '========================================
                     Tester.Print "next one press"
                      CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                    
                      Call MsecDelay(0.2)
                      CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                  
                      Tester.Print "next one up"
                        
                        
                         
                      '=========================================
                      '
                      '   End Next one song WMA
                      '
                      '========================================
                    
                     ' Call MsecDelay(1#)
                     '  CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                       Call DACGetData2ChannelUpTriggle
                       Call ADCBuffer2AudioBuffer(1)
                       Call CompareDataTrace(1)
                           
                       rv3 = WMACompareResult
                 Else
                   rv3 = 4
                 End If
                 Call LabelMenu(3, rv3, rv2)
                  
                 '========================================================
                 '
                 '  Recode Raw data to pattern file
                 '
                 '========================================================
                  
                  
                 If Tester.AudioRecordCheck.Value = 1 Then
                                       
                     Open App.Path & "\MP3Pattern\MP3.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, MP3RawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, MP3RawData(1, i)
                         Next i
                      
    
                    Close #4
                
                
                                  
                     Open App.Path & "\WMAPattern\WMA.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, WMARawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, WMARawData(1, i)
                         Next i
                      
    
                    Close #4
                    
                 End If
                 
                 If rv2 * rv3 = 1 Then
                 
                   ' Tester.Cls
                    Tester.Print "MP3Matchmode="; MP3MatchIndex
                    Tester.Print "WMAMatchmode="; WMAMatchIndex
                 End If

                 dma2205_4.Show
                 dma2205_4.ShowData 0, 6000
                
                 Card2214Exist = 0
                  D2K_Release_Card (Card2214)
                  CardResult = DO_WritePort(card, Channel_P1A, &HC0)  ' Allen debug
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
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
                ElseIf rv0 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub
Public Sub AU3152ALF2C_1_SlotTest_Noise()
Dim i As Integer
  
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                 '=========================================
                '    POWER on
                '=========================================
               '   CardResult = DO_WritePort(card, Channel_P1A, &H80)
                    CardResult = DO_WritePort(card, Channel_P1A, &HC0)  ' Allen debug
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                CardResult = DO_WritePort(card, Channel_P1A, &H6F)
              '    CardResult = DO_WritePort(card, Channel_P1A, &H2F)   ' usb Power Off
                  Call MsecDelay(0.8)    'power on time
                  
              '   Call MsecDelay(1#)      'power on time
                  
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
            
                ClosePipe
                LBA = LBA + 1
                OldLBa = LBA
       
                
                LBA = LBA + 40000 ' 20000 mp3 shift
                
                rv0 = CBWTest_New_31303(0, 1, "vid_058f")
                LBA = OldLBa
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 Tester.Cls
                Tester.Print "LBA="; LBA
                Tester.Print "rv0="; rv0
                 
                If rv0 = 1 Then
                    
                          '=================================
                          '   begin load pattern
                          '=================================
                          
                         
                          If OldChipName <> ChipName Then
                                 
                              '  Call LoadAU3150JLF27AudioPattern
                                OldChipName = ChipName
                          End If
                                                      
              
                          
                          '==========================================
                          '  do a low-high-low pulse at play button
                          '==========================================
                         
                             
                             
                             CardResult = DO_WritePort(card, Channel_P1A, &HCF)
                             
                             Call MsecDelay(0.3)
                             CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                                 Call MsecDelay(2.5)
                                'Call MsecDelay(4.5)
                              CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                            '   CardResult = DO_WritePort(card, Channel_P1A, &HDF)
                          '    Call MsecDelay(4)    'power on time and DAC initial time
                               
                            ' Call MsecDelay(2.5)     'for 3150K power on time
                               Call MsecDelay(0.3)
                             CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                            
                            ' CardResult = DO_WritePort(card, Channel_P1A, &HCF)
                          '==========================================
                          '  Get ADC Data and compare result
                          '==========================================
                   Call MsecDelay(1.5)
                           Call DACGetData2ChannelUpTriggle
                           Call ADCBuffer2AudioBuffer(0)
                           Call DistingDataTrace(0)
                           
                           rv2 = MP3DistinguishResult
                  Else
                   rv2 = 4
                 End If
                 
                 Call LabelMenu(2, rv2, rv0)
                 

               
                If (rv0 = 1) And (rv2 = 1 Or Tester.AudioRecordCheck.Value = 1) Then
                  
                      D2K_Release_Card (Card2214)
                     Card2214Exist = 0
                     
                    '=========================================
                    '
                    '   Start Next one song WMA
                    '
                    '========================================
                     Tester.Print "next one press"
                      CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                    
                      Call MsecDelay(0.2)
                      CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                  
                      Tester.Print "next one up"
                        
                        
                         
                      '=========================================
                      '
                      '   End Next one song WMA
                      '
                      '========================================
                    
                       Call MsecDelay(0.5)
                     '  CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                       Call DACGetData2ChannelUpTriggle
                       Call ADCBuffer2AudioBuffer(1)
                            Call DistingDataTrace(1)
                           
                           rv3 = WMADistinguishResult
                 Else
                   rv3 = 4
                 End If
                 Call LabelMenu(3, rv3, rv2)
                  
                 '========================================================
                 '
                 '  Recode Raw data to pattern file
                 '
                 '========================================================
                  
                  
                 If Tester.AudioRecordCheck.Value = 1 Then
                                       
                     Open App.Path & "\MP3Pattern\MP3.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, MP3RawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, MP3RawData(1, i)
                         Next i
                      
    
                    Close #4
                
                
                                  
                     Open App.Path & "\WMAPattern\WMA.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, WMARawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, WMARawData(1, i)
                         Next i
                      
    
                    Close #4
                    
                 End If
                 
                 If rv2 * rv3 = 1 Then
                 
                   ' Tester.Cls
                    Tester.Print "MP3Matchmode="; MP3MatchIndex
                    Tester.Print "WMAMatchmode="; WMAMatchIndex
                 End If

                 dma2205_4.Show
                 dma2205_4.ShowData3 0, 6000
                
                 Card2214Exist = 0
                  D2K_Release_Card (Card2214)
                  CardResult = DO_WritePort(card, Channel_P1A, &HC0)  ' Allen debug
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
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
                ElseIf rv0 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub



Public Sub AU3152ALF2D_1_SlotTest_Noise()
Dim i As Integer
  
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                 '=========================================
                '    POWER on
                '=========================================
               '   CardResult = DO_WritePort(card, Channel_P1A, &H80)
                    CardResult = DO_WritePort(card, Channel_P1A, &HC0)  ' Allen debug
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                CardResult = DO_WritePort(card, Channel_P1A, &H6F)
              '    CardResult = DO_WritePort(card, Channel_P1A, &H2F)   ' usb Power Off
                  Call MsecDelay(0.8)    'power on time
                  
              '   Call MsecDelay(1#)      'power on time
                  
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
            
                ClosePipe
                LBA = LBA + 1
                OldLBa = LBA
       
                
                LBA = LBA + 40000 ' 20000 mp3 shift
                
                rv0 = CBWTest_New_31303(0, 1, "vid_058f")
                LBA = OldLBa
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 Tester.Cls
                Tester.Print "LBA="; LBA
                Tester.Print "rv0="; rv0
                 
                If rv0 = 1 Then
                    
                          '=================================
                          '   begin load pattern
                          '=================================
                          
                         
                          If OldChipName <> ChipName Then
                                 
                              '  Call LoadAU3150JLF27AudioPattern
                                OldChipName = ChipName
                          End If
                                                      
              
                          
                          '==========================================
                          '  do a low-high-low pulse at play button
                          '==========================================
                         
                             
                             
                             CardResult = DO_WritePort(card, Channel_P1A, &HCF)
                             
                             Call MsecDelay(0.3)
                             CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                                 Call MsecDelay(2.5)
                                'Call MsecDelay(4.5)
                              CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                            '   CardResult = DO_WritePort(card, Channel_P1A, &HDF)
                          '    Call MsecDelay(4)    'power on time and DAC initial time
                               
                            ' Call MsecDelay(2.5)     'for 3150K power on time
                               Call MsecDelay(0.3)
                             CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                            
                            ' CardResult = DO_WritePort(card, Channel_P1A, &HCF)
                          '==========================================
                          '  Get ADC Data and compare result
                          '==========================================
                   Call MsecDelay(1.5)
                           Call DACGetData2ChannelUpTriggle
                           Call ADCBuffer2AudioBuffer(0)
                           Call DistingDataTrace(0)
                           
                           rv2 = MP3DistinguishResult
                  Else
                   rv2 = 4
                 End If
                 
                 Call LabelMenu(2, rv2, rv0)
                 

               
                If (rv0 = 1) And (rv2 = 1 Or Tester.AudioRecordCheck.Value = 1) Then
                  
                      D2K_Release_Card (Card2214)
                     Card2214Exist = 0
                     
                    '=========================================
                    '
                    '   Start Next one song WMA
                    '
                    '========================================
                     Tester.Print "next one press"
                      CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                    
                      Call MsecDelay(0.2)
                      CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                  
                      Tester.Print "next one up"
                        
                        
                         
                      '=========================================
                      '
                      '   End Next one song WMA
                      '
                      '========================================
                    
                       Call MsecDelay(0.5)
                     '  CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                       Call DACGetData2ChannelUpTriggle
                       Call ADCBuffer2AudioBuffer(1)
                            Call DistingDataTrace(1)
                           
                           rv3 = WMADistinguishResult
                 Else
                   rv3 = 4
                 End If
                 Call LabelMenu(3, rv3, rv2)
                  
                 '========================================================
                 '
                 '  Recode Raw data to pattern file
                 '
                 '========================================================
                  
                  
                 If Tester.AudioRecordCheck.Value = 1 Then
                                       
                     Open App.Path & "\MP3Pattern\MP3.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, MP3RawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, MP3RawData(1, i)
                         Next i
                      
    
                    Close #4
                
                
                                  
                     Open App.Path & "\WMAPattern\WMA.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, WMARawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, WMARawData(1, i)
                         Next i
                      
    
                    Close #4
                    
                 End If
                 
                 If rv2 * rv3 = 1 Then
                 
                   ' Tester.Cls
                    Tester.Print "MP3Matchmode="; MP3MatchIndex
                    Tester.Print "WMAMatchmode="; WMAMatchIndex
                 End If

                 dma2205_4.Show
                 dma2205_4.ShowData3 0, 6000
                
                 Card2214Exist = 0
                  D2K_Release_Card (Card2214)
                  CardResult = DO_WritePort(card, Channel_P1A, &HC0)  ' Allen debug
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
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
                ElseIf rv0 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub

Public Sub AU3152ALF28_1_SlotTest_Noise()
Dim i As Integer
  
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                 '=========================================
                '    POWER on
                '=========================================
               '   CardResult = DO_WritePort(card, Channel_P1A, &H80)
                    CardResult = DO_WritePort(card, Channel_P1A, &HC0)  ' Allen debug
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                CardResult = DO_WritePort(card, Channel_P1A, &H6F)
              '    CardResult = DO_WritePort(card, Channel_P1A, &H2F)   ' usb Power Off
                  Call MsecDelay(0.8)    'power on time
                  
              '   Call MsecDelay(1#)      'power on time
                  
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
            
                ClosePipe
                LBA = LBA + 1
                OldLBa = LBA
       
                
                LBA = LBA + 40000 ' 20000 mp3 shift
                
                rv0 = CBWTest_New_31303(0, 1, "vid_058f")
                LBA = OldLBa
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 Tester.Cls
                Tester.Print "LBA="; LBA
                Tester.Print "rv0="; rv0
                 
                If rv0 = 1 Then
                    
                          '=================================
                          '   begin load pattern
                          '=================================
                          
                         
                          If OldChipName <> ChipName Then
                                 
                                Call LoadAU3150JLF27AudioPattern
                                OldChipName = ChipName
                          End If
                                                      
              
                          
                          '==========================================
                          '  do a low-high-low pulse at play button
                          '==========================================
                         
                             
                             
                             CardResult = DO_WritePort(card, Channel_P1A, &HCF)
                             
                             Call MsecDelay(0.3)
                             CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                                 Call MsecDelay(2.5)
                                'Call MsecDelay(4.5)
                              CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                            '   CardResult = DO_WritePort(card, Channel_P1A, &HDF)
                          '    Call MsecDelay(4)    'power on time and DAC initial time
                               
                            ' Call MsecDelay(2.5)     'for 3150K power on time
                               Call MsecDelay(0.3)
                             CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                            
                            ' CardResult = DO_WritePort(card, Channel_P1A, &HCF)
                          '==========================================
                          '  Get ADC Data and compare result
                          '==========================================
                   Call MsecDelay(1.5)
                           Call DACGetData2ChannelUpTriggle
                           Call ADCBuffer2AudioBuffer(0)
                           Call CompareDataTrace(0)
                           
                           rv2 = MP3CompareResult
                  Else
                   rv2 = 4
                 End If
                 
                 Call LabelMenu(2, rv2, rv0)
                 

               
                If (rv0 = 1) And (rv2 = 1 Or Tester.AudioRecordCheck.Value = 1) Then
                  
                      D2K_Release_Card (Card2214)
                     Card2214Exist = 0
                     
                    '=========================================
                    '
                    '   Start Next one song WMA
                    '
                    '========================================
                     Tester.Print "next one press"
                      CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                    
                      Call MsecDelay(0.2)
                      CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                  
                      Tester.Print "next one up"
                        
                        
                         
                      '=========================================
                      '
                      '   End Next one song WMA
                      '
                      '========================================
                    
                       Call MsecDelay(0.5)
                     '  CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                       Call DACGetData2ChannelUpTriggle
                       Call ADCBuffer2AudioBuffer(1)
                       Call CompareDataTrace(1)
                           
                       rv3 = WMACompareResult
                 Else
                   rv3 = 4
                 End If
                 Call LabelMenu(3, rv3, rv2)
                  
                 '========================================================
                 '
                 '  Recode Raw data to pattern file
                 '
                 '========================================================
                  
                  
                 If Tester.AudioRecordCheck.Value = 1 Then
                                       
                     Open App.Path & "\MP3Pattern\MP3.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, MP3RawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, MP3RawData(1, i)
                         Next i
                      
    
                    Close #4
                
                
                                  
                     Open App.Path & "\WMAPattern\WMA.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, WMARawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, WMARawData(1, i)
                         Next i
                      
    
                    Close #4
                    
                 End If
                 
                 If rv2 * rv3 = 1 Then
                 
                   ' Tester.Cls
                    Tester.Print "MP3Matchmode="; MP3MatchIndex
                    Tester.Print "WMAMatchmode="; WMAMatchIndex
                 End If

                 dma2205_4.Show
                 dma2205_4.ShowData 0, 6000
                
                 Card2214Exist = 0
                  D2K_Release_Card (Card2214)
                  CardResult = DO_WritePort(card, Channel_P1A, &HC0)  ' Allen debug
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
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
                ElseIf rv0 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub

Public Sub AU3150PLF28_1_SlotTest_Noise()
Dim i As Integer
  
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                 '=========================================
                '    POWER on
                '=========================================
               '   CardResult = DO_WritePort(card, Channel_P1A, &H80)
                    CardResult = DO_WritePort(card, Channel_P1A, &HC0)  ' Allen debug
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                CardResult = DO_WritePort(card, Channel_P1A, &H6F)
              '    CardResult = DO_WritePort(card, Channel_P1A, &H2F)   ' usb Power Off
                  Call MsecDelay(0.8)    'power on time
                  
              '   Call MsecDelay(1#)      'power on time
                  
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
            
                ClosePipe
                LBA = LBA + 1
                OldLBa = LBA
       
                
                LBA = LBA + 40000 ' 20000 mp3 shift
                
                rv0 = CBWTest_New_31303(0, 1, "vid_058f")
                LBA = OldLBa
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 Tester.Cls
                Tester.Print "LBA="; LBA
                Tester.Print "rv0="; rv0
                 
                If rv0 = 1 Then
                    
                          '=================================
                          '   begin load pattern
                          '=================================
                          
                         
                          If OldChipName <> ChipName Then
                                 
                                Call LoadAU3150JLF27AudioPattern
                                OldChipName = ChipName
                          End If
                                                      
              
                          
                          '==========================================
                          '  do a low-high-low pulse at play button
                          '==========================================
                         
                             
                             
                             CardResult = DO_WritePort(card, Channel_P1A, &HCF)
                             
                             Call MsecDelay(0.3)
                             CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                                 Call MsecDelay(2.5)
                                'Call MsecDelay(4.5)
                              CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                            '   CardResult = DO_WritePort(card, Channel_P1A, &HDF)
                          '    Call MsecDelay(4)    'power on time and DAC initial time
                               
                            ' Call MsecDelay(2.5)     'for 3150K power on time
                               Call MsecDelay(0.3)
                             CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                            
                            ' CardResult = DO_WritePort(card, Channel_P1A, &HCF)
                          '==========================================
                          '  Get ADC Data and compare result
                          '==========================================
                   Call MsecDelay(1.5)
                           Call DACGetData2ChannelUpTriggle
                           Call ADCBuffer2AudioBuffer(0)
                           Call CompareDataTrace(0)
                           
                           rv2 = MP3CompareResult
                  Else
                   rv2 = 4
                 End If
                 
                 Call LabelMenu(2, rv2, rv0)
                 

               
                If (rv0 = 1) And (rv2 = 1 Or Tester.AudioRecordCheck.Value = 1) Then
                  
                      D2K_Release_Card (Card2214)
                     Card2214Exist = 0
                     
                    '=========================================
                    '
                    '   Start Next one song WMA
                    '
                    '========================================
                     Tester.Print "next one press"
                      CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                    
                      Call MsecDelay(0.2)
                      CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                  
                      Tester.Print "next one up"
                        
                        
                         
                      '=========================================
                      '
                      '   End Next one song WMA
                      '
                      '========================================
                    
                     ' Call MsecDelay(1#)
                     '  CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                       Call DACGetData2ChannelUpTriggle
                       Call ADCBuffer2AudioBuffer(1)
                       Call CompareDataTrace(1)
                           
                       rv3 = WMACompareResult
                 Else
                   rv3 = 4
                 End If
                 Call LabelMenu(3, rv3, rv2)
                  
                 '========================================================
                 '
                 '  Recode Raw data to pattern file
                 '
                 '========================================================
                  
                  
                 If Tester.AudioRecordCheck.Value = 1 Then
                                       
                     Open App.Path & "\MP3Pattern\MP3.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, MP3RawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, MP3RawData(1, i)
                         Next i
                      
    
                    Close #4
                
                
                                  
                     Open App.Path & "\WMAPattern\WMA.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, WMARawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, WMARawData(1, i)
                         Next i
                      
    
                    Close #4
                    
                 End If
                 
                 If rv2 * rv3 = 1 Then
                 
                   ' Tester.Cls
                    Tester.Print "MP3Matchmode="; MP3MatchIndex
                    Tester.Print "WMAMatchmode="; WMAMatchIndex
                 End If

                 dma2205_4.Show
                 dma2205_4.ShowData 0, 6000
                
                 Card2214Exist = 0
                  D2K_Release_Card (Card2214)
                  CardResult = DO_WritePort(card, Channel_P1A, &HC0)  ' Allen debug
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
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
                ElseIf rv0 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub
Public Sub AU3150NLF28_1_SlotTest_Noise()
Dim i As Integer
  
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                 '=========================================
                '    POWER on
                '=========================================
               '   CardResult = DO_WritePort(card, Channel_P1A, &H80)
                    CardResult = DO_WritePort(card, Channel_P1A, &HC0)  ' Allen debug
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                CardResult = DO_WritePort(card, Channel_P1A, &H6F)
              '    CardResult = DO_WritePort(card, Channel_P1A, &H2F)   ' usb Power Off
                  Call MsecDelay(0.8)    'power on time
                  
              '   Call MsecDelay(1#)      'power on time
                  
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
            
                ClosePipe
                LBA = LBA + 1
                OldLBa = LBA
       
                
                LBA = LBA + 40000 ' 20000 mp3 shift
                
                rv0 = CBWTest_New_31303(0, 1, "vid_058f")
                LBA = OldLBa
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 Tester.Cls
                Tester.Print "LBA="; LBA
                Tester.Print "rv0="; rv0
                 
                If rv0 = 1 Then
                    
                          '=================================
                          '   begin load pattern
                          '=================================
                          
                         
                          If OldChipName <> ChipName Then
                                 
                                Call LoadAU3150JLF27AudioPattern
                                OldChipName = ChipName
                          End If
                                                      
              
                          
                          '==========================================
                          '  do a low-high-low pulse at play button
                          '==========================================
                         
                             
                             
                             CardResult = DO_WritePort(card, Channel_P1A, &HCF)
                             
                             Call MsecDelay(0.3)
                             CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                                 Call MsecDelay(2.5)
                                'Call MsecDelay(4.5)
                              CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                            '   CardResult = DO_WritePort(card, Channel_P1A, &HDF)
                          '    Call MsecDelay(4)    'power on time and DAC initial time
                               
                            ' Call MsecDelay(2.5)     'for 3150K power on time
                               Call MsecDelay(0.3)
                             CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                            
                            ' CardResult = DO_WritePort(card, Channel_P1A, &HCF)
                          '==========================================
                          '  Get ADC Data and compare result
                          '==========================================
                   Call MsecDelay(1.5)
                           Call DACGetData2ChannelUpTriggle
                           Call ADCBuffer2AudioBuffer(0)
                           Call CompareDataTrace(0)
                           
                           rv2 = MP3CompareResult
                  Else
                   rv2 = 4
                 End If
                 
                 Call LabelMenu(2, rv2, rv0)
                 

               
                If (rv0 = 1) And (rv2 = 1 Or Tester.AudioRecordCheck.Value = 1) Then
                  
                      D2K_Release_Card (Card2214)
                     Card2214Exist = 0
                     
                    '=========================================
                    '
                    '   Start Next one song WMA
                    '
                    '========================================
                     Tester.Print "next one press"
                      CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                    
                      Call MsecDelay(0.2)
                      CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                  
                      Tester.Print "next one up"
                        
                        
                         
                      '=========================================
                      '
                      '   End Next one song WMA
                      '
                      '========================================
                    
                     ' Call MsecDelay(1#)
                     '  CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                       Call DACGetData2ChannelUpTriggle
                       Call ADCBuffer2AudioBuffer(1)
                       Call CompareDataTrace(1)
                           
                       rv3 = WMACompareResult
                 Else
                   rv3 = 4
                 End If
                 Call LabelMenu(3, rv3, rv2)
                  
                 '========================================================
                 '
                 '  Recode Raw data to pattern file
                 '
                 '========================================================
                  
                  
                 If Tester.AudioRecordCheck.Value = 1 Then
                                       
                     Open App.Path & "\MP3Pattern\MP3.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, MP3RawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, MP3RawData(1, i)
                         Next i
                      
    
                    Close #4
                
                
                                  
                     Open App.Path & "\WMAPattern\WMA.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, WMARawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, WMARawData(1, i)
                         Next i
                      
    
                    Close #4
                    
                 End If
                 
                 If rv2 * rv3 = 1 Then
                 
                   ' Tester.Cls
                    Tester.Print "MP3Matchmode="; MP3MatchIndex
                    Tester.Print "WMAMatchmode="; WMAMatchIndex
                 End If

                 dma2205_4.Show
                 dma2205_4.ShowData 0, 6000
                
                 Card2214Exist = 0
                  D2K_Release_Card (Card2214)
                  CardResult = DO_WritePort(card, Channel_P1A, &HC0)  ' Allen debug
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
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
                ElseIf rv0 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub

Public Sub AU3150LLF28_1_SlotTest_Noise()
Dim i As Integer
  
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                 '=========================================
                '    POWER on
                '=========================================
               '   CardResult = DO_WritePort(card, Channel_P1A, &H80)
                    CardResult = DO_WritePort(card, Channel_P1A, &HC0)  ' Allen debug
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                CardResult = DO_WritePort(card, Channel_P1A, &H6F)
              '    CardResult = DO_WritePort(card, Channel_P1A, &H2F)   ' usb Power Off
                  Call MsecDelay(0.8)    'power on time
                  
              '   Call MsecDelay(1#)      'power on time
                  
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
            
                ClosePipe
                LBA = LBA + 1
                OldLBa = LBA
       
                
                LBA = LBA + 40000 ' 20000 mp3 shift
                
                rv0 = CBWTest_New_31303(0, 1, "vid_058f")
                LBA = OldLBa
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 Tester.Cls
                Tester.Print "LBA="; LBA
                Tester.Print "rv0="; rv0
                 
                If rv0 = 1 Then
                    
                          '=================================
                          '   begin load pattern
                          '=================================
                          
                         
                          If OldChipName <> ChipName Then
                                 
                                Call LoadAU3150JLF27AudioPattern
                                OldChipName = ChipName
                          End If
                                                      
              
                          
                          '==========================================
                          '  do a low-high-low pulse at play button
                          '==========================================
                         
                             
                             
                             CardResult = DO_WritePort(card, Channel_P1A, &HCF)
                             
                             Call MsecDelay(0.3)
                             CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                                 Call MsecDelay(2.5)
                                'Call MsecDelay(4.5)
                              CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                            '   CardResult = DO_WritePort(card, Channel_P1A, &HDF)
                          '    Call MsecDelay(4)    'power on time and DAC initial time
                               
                            ' Call MsecDelay(2.5)     'for 3150K power on time
                               Call MsecDelay(0.3)
                             CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                            
                            ' CardResult = DO_WritePort(card, Channel_P1A, &HCF)
                          '==========================================
                          '  Get ADC Data and compare result
                          '==========================================
                   Call MsecDelay(1.5)
                           Call DACGetData2ChannelUpTriggle
                           Call ADCBuffer2AudioBuffer(0)
                           Call CompareDataTrace(0)
                           
                           rv2 = MP3CompareResult
                  Else
                   rv2 = 4
                 End If
                 
                 Call LabelMenu(2, rv2, rv0)
                 

               
                If (rv0 = 1) And (rv2 = 1 Or Tester.AudioRecordCheck.Value = 1) Then
                  
                      D2K_Release_Card (Card2214)
                     Card2214Exist = 0
                     
                    '=========================================
                    '
                    '   Start Next one song WMA
                    '
                    '========================================
                     Tester.Print "next one press"
                      CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                    
                      Call MsecDelay(0.2)
                      CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                  
                      Tester.Print "next one up"
                        
                        
                         
                      '=========================================
                      '
                      '   End Next one song WMA
                      '
                      '========================================
                    
                     ' Call MsecDelay(1#)
                     '  CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                       Call DACGetData2ChannelUpTriggle
                       Call ADCBuffer2AudioBuffer(1)
                       Call CompareDataTrace(1)
                           
                       rv3 = WMACompareResult
                 Else
                   rv3 = 4
                 End If
                 Call LabelMenu(3, rv3, rv2)
                  
                 '========================================================
                 '
                 '  Recode Raw data to pattern file
                 '
                 '========================================================
                  
                  
                 If Tester.AudioRecordCheck.Value = 1 Then
                                       
                     Open App.Path & "\MP3Pattern\MP3.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, MP3RawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, MP3RawData(1, i)
                         Next i
                      
    
                    Close #4
                
                
                                  
                     Open App.Path & "\WMAPattern\WMA.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, WMARawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, WMARawData(1, i)
                         Next i
                      
    
                    Close #4
                    
                 End If
                 
                 If rv2 * rv3 = 1 Then
                 
                   ' Tester.Cls
                    Tester.Print "MP3Matchmode="; MP3MatchIndex
                    Tester.Print "WMAMatchmode="; WMAMatchIndex
                 End If

                 dma2205_4.Show
                 dma2205_4.ShowData 0, 6000
                
                 Card2214Exist = 0
                  D2K_Release_Card (Card2214)
                  CardResult = DO_WritePort(card, Channel_P1A, &HC0)  ' Allen debug
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
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
                ElseIf rv0 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub

Public Sub AU3150LLF2C_1_SlotTest_Noise()
Dim i As Integer
  
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                 '=========================================
                '    POWER on
                '=========================================
               '   CardResult = DO_WritePort(card, Channel_P1A, &H80)
                    CardResult = DO_WritePort(card, Channel_P1A, &HC0)  ' Allen debug
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                CardResult = DO_WritePort(card, Channel_P1A, &H6F)
              '    CardResult = DO_WritePort(card, Channel_P1A, &H2F)   ' usb Power Off
                  Call MsecDelay(0.8)    'power on time
                  
              '   Call MsecDelay(1#)      'power on time
                  
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
            
                ClosePipe
                LBA = LBA + 1
                OldLBa = LBA
       
                
                LBA = LBA + 40000 ' 20000 mp3 shift
                
                rv0 = CBWTest_New_31303(0, 1, "vid_058f")
                LBA = OldLBa
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 Tester.Cls
                Tester.Print "LBA="; LBA
                Tester.Print "rv0="; rv0
                 
                If rv0 = 1 Then
                    
                          '=================================
                          '   begin load pattern
                          '=================================
                          
                         
                          If OldChipName <> ChipName Then
                                 
                           '     Call LoadAU3150JLF27AudioPattern
                                OldChipName = ChipName
                          End If
                                                      
              
                          
                          '==========================================
                          '  do a low-high-low pulse at play button
                          '==========================================
                         
                             
                             
                             CardResult = DO_WritePort(card, Channel_P1A, &HCF)
                             
                             Call MsecDelay(0.3)
                             CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                                 Call MsecDelay(2.5)
                                'Call MsecDelay(4.5)
                              CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                            '   CardResult = DO_WritePort(card, Channel_P1A, &HDF)
                          '    Call MsecDelay(4)    'power on time and DAC initial time
                               
                            ' Call MsecDelay(2.5)     'for 3150K power on time
                               Call MsecDelay(0.3)
                             CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                            
                            ' CardResult = DO_WritePort(card, Channel_P1A, &HCF)
                          '==========================================
                          '  Get ADC Data and compare result
                          '==========================================
                   Call MsecDelay(1.5)
                           Call DACGetData2ChannelUpTriggle
                           Call ADCBuffer2AudioBuffer(0)
                           
                             Call DistingDataTrace(0)
                            rv2 = MP3DistinguishResult
                       '    Call CompareDataTrace(0)
                           
                       '    rv2 = MP3CompareResult
                  Else
                   rv2 = 4
                 End If
                 
                 Call LabelMenu(2, rv2, rv0)
                 

               
                If (rv0 = 1) And (rv2 = 1 Or Tester.AudioRecordCheck.Value = 1) Then
                  
                      D2K_Release_Card (Card2214)
                     Card2214Exist = 0
                     
                    '=========================================
                    '
                    '   Start Next one song WMA
                    '
                    '========================================
                     Tester.Print "next one press"
                      CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                    
                      Call MsecDelay(0.2)
                      CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                  
                      Tester.Print "next one up"
                        
                        
                         
                      '=========================================
                      '
                      '   End Next one song WMA
                      '
                      '========================================
                    
                     ' Call MsecDelay(1#)
                     '  CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                       Call DACGetData2ChannelUpTriggle
                       Call ADCBuffer2AudioBuffer(1)
                         Call DistingDataTrace(1)
                         rv3 = WMADistinguishResult
                   '    Call CompareDataTrace(1)
                           
                   '    rv3 = WMACompareResult
                 Else
                   rv3 = 4
                 End If
                 Call LabelMenu(3, rv3, rv2)
                  
                 '========================================================
                 '
                 '  Recode Raw data to pattern file
                 '
                 '========================================================
                  
                  
                 If Tester.AudioRecordCheck.Value = 1 Then
                                       
                     Open App.Path & "\MP3Pattern\MP3.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, MP3RawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, MP3RawData(1, i)
                         Next i
                      
    
                    Close #4
                
                
                                  
                     Open App.Path & "\WMAPattern\WMA.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, WMARawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, WMARawData(1, i)
                         Next i
                      
    
                    Close #4
                    
                 End If
                 
                 If rv2 * rv3 = 1 Then
                 
                   ' Tester.Cls
                    Tester.Print "MP3Matchmode="; MP3MatchIndex
                    Tester.Print "WMAMatchmode="; WMAMatchIndex
                 End If

                 dma2205_4.Show
                 dma2205_4.ShowData3 0, 6000
                
                 Card2214Exist = 0
                  D2K_Release_Card (Card2214)
                  CardResult = DO_WritePort(card, Channel_P1A, &HC0)  ' Allen debug
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
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
                ElseIf rv0 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub

Public Sub AU3150LLF2D_1_SlotTest_Noise()
Dim i As Integer
  
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                 '=========================================
                '    POWER on
                '=========================================
               '   CardResult = DO_WritePort(card, Channel_P1A, &H80)
                    CardResult = DO_WritePort(card, Channel_P1A, &HC0)  ' Allen debug
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                CardResult = DO_WritePort(card, Channel_P1A, &H6F)
              '    CardResult = DO_WritePort(card, Channel_P1A, &H2F)   ' usb Power Off
                  Call MsecDelay(0.8)    'power on time
                  
              '   Call MsecDelay(1#)      'power on time
                  
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
            
                ClosePipe
                LBA = LBA + 1
                OldLBa = LBA
       
                
                LBA = LBA + 40000 ' 20000 mp3 shift
                
                rv0 = CBWTest_New_31303(0, 1, "vid_058f")
                LBA = OldLBa
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 Tester.Cls
                Tester.Print "LBA="; LBA
                Tester.Print "rv0="; rv0
                 
                If rv0 = 1 Then
                    
                          '=================================
                          '   begin load pattern
                          '=================================
                          
                         
                          If OldChipName <> ChipName Then
                                 
                           '     Call LoadAU3150JLF27AudioPattern
                                OldChipName = ChipName
                          End If
                                                      
              
                          
                          '==========================================
                          '  do a low-high-low pulse at play button
                          '==========================================
                         
                             
                             
                             CardResult = DO_WritePort(card, Channel_P1A, &HCF)
                             
                             Call MsecDelay(0.3)
                             CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                                 Call MsecDelay(2.5)
                                'Call MsecDelay(4.5)
                              CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                            '   CardResult = DO_WritePort(card, Channel_P1A, &HDF)
                          '    Call MsecDelay(4)    'power on time and DAC initial time
                               
                            ' Call MsecDelay(2.5)     'for 3150K power on time
                               Call MsecDelay(0.3)
                             CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                            
                            ' CardResult = DO_WritePort(card, Channel_P1A, &HCF)
                          '==========================================
                          '  Get ADC Data and compare result
                          '==========================================
                   Call MsecDelay(1.5)
                           Call DACGetData2ChannelUpTriggle
                           Call ADCBuffer2AudioBuffer(0)
                           
                             Call DistingDataTrace(0)
                            rv2 = MP3DistinguishResult
                       '    Call CompareDataTrace(0)
                           
                       '    rv2 = MP3CompareResult
                  Else
                   rv2 = 4
                 End If
                 
                 Call LabelMenu(2, rv2, rv0)
                 

               
                If (rv0 = 1) And (rv2 = 1 Or Tester.AudioRecordCheck.Value = 1) Then
                  
                      D2K_Release_Card (Card2214)
                     Card2214Exist = 0
                     
                    '=========================================
                    '
                    '   Start Next one song WMA
                    '
                    '========================================
                     Tester.Print "next one press"
                      CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                    
                      Call MsecDelay(0.2)
                      CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                  
                      Tester.Print "next one up"
                        
                        
                         
                      '=========================================
                      '
                      '   End Next one song WMA
                      '
                      '========================================
                    
                     ' Call MsecDelay(1#)
                     '  CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                       Call DACGetData2ChannelUpTriggle
                       Call ADCBuffer2AudioBuffer(1)
                         Call DistingDataTrace(1)
                         rv3 = WMADistinguishResult
                   '    Call CompareDataTrace(1)
                           
                   '    rv3 = WMACompareResult
                 Else
                   rv3 = 4
                 End If
                 Call LabelMenu(3, rv3, rv2)
                  
                 '========================================================
                 '
                 '  Recode Raw data to pattern file
                 '
                 '========================================================
                  
                  
                 If Tester.AudioRecordCheck.Value = 1 Then
                                       
                     Open App.Path & "\MP3Pattern\MP3.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, MP3RawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, MP3RawData(1, i)
                         Next i
                      
    
                    Close #4
                
                
                                  
                     Open App.Path & "\WMAPattern\WMA.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, WMARawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, WMARawData(1, i)
                         Next i
                      
    
                    Close #4
                    
                 End If
                 
                 If rv2 * rv3 = 1 Then
                 
                   ' Tester.Cls
                    Tester.Print "MP3Matchmode="; MP3MatchIndex
                    Tester.Print "WMAMatchmode="; WMAMatchIndex
                 End If

                 dma2205_4.Show
                 dma2205_4.ShowData3 0, 6000
                
                 Card2214Exist = 0
                  D2K_Release_Card (Card2214)
                  CardResult = DO_WritePort(card, Channel_P1A, &HC0)  ' Allen debug
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
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
                ElseIf rv0 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub

Public Sub AU3152CLF28_1_SlotTest_Noise()
Dim i As Integer
  
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                 '=========================================
                '    POWER on
                '=========================================
               '   CardResult = DO_WritePort(card, Channel_P1A, &H80)
                    CardResult = DO_WritePort(card, Channel_P1A, &HC0)  ' Allen debug
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                CardResult = DO_WritePort(card, Channel_P1A, &H6F)
              '    CardResult = DO_WritePort(card, Channel_P1A, &H2F)   ' usb Power Off
                  Call MsecDelay(0.8)    'power on time
                  
              '   Call MsecDelay(1#)      'power on time
                  
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
            
                ClosePipe
                LBA = LBA + 1
                OldLBa = LBA
       
                
                LBA = LBA + 40000 ' 20000 mp3 shift
                
                rv0 = CBWTest_New_31303(0, 1, "vid_058f")
                LBA = OldLBa
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 Tester.Cls
                Tester.Print "LBA="; LBA
                Tester.Print "rv0="; rv0
                 
                If rv0 = 1 Then
                    
                          '=================================
                          '   begin load pattern
                          '=================================
                          
                         
                          If OldChipName <> ChipName Then
                                 
                                Call LoadAU3150JLF27AudioPattern
                                OldChipName = ChipName
                          End If
                                                      
              
                          
                          '==========================================
                          '  do a low-high-low pulse at play button
                          '==========================================
                         
                             
                             
                             CardResult = DO_WritePort(card, Channel_P1A, &HCF)
                             
                             Call MsecDelay(0.3)
                             CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                                 Call MsecDelay(2.5)
                                'Call MsecDelay(4.5)
                              CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                            '   CardResult = DO_WritePort(card, Channel_P1A, &HDF)
                          '    Call MsecDelay(4)    'power on time and DAC initial time
                               
                            ' Call MsecDelay(2.5)     'for 3150K power on time
                               Call MsecDelay(0.3)
                             CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                            
                            ' CardResult = DO_WritePort(card, Channel_P1A, &HCF)
                          '==========================================
                          '  Get ADC Data and compare result
                          '==========================================
                   Call MsecDelay(1.5)
                           Call DACGetData2ChannelUpTriggle
                           Call ADCBuffer2AudioBuffer(0)
                           Call CompareDataTrace(0)
                           
                           rv2 = MP3CompareResult
                  Else
                   rv2 = 4
                 End If
                 
                 Call LabelMenu(2, rv2, rv0)
                 

               
                If (rv0 = 1) And (rv2 = 1 Or Tester.AudioRecordCheck.Value = 1) Then
                  
                      D2K_Release_Card (Card2214)
                     Card2214Exist = 0
                     
                    '=========================================
                    '
                    '   Start Next one song WMA
                    '
                    '========================================
                     Tester.Print "next one press"
                      CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                    
                      Call MsecDelay(0.2)
                      CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                  
                      Tester.Print "next one up"
                        
                        
                         
                      '=========================================
                      '
                      '   End Next one song WMA
                      '
                      '========================================
                    
                       Call MsecDelay(0.5)
                     '  CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                       Call DACGetData2ChannelUpTriggle
                       Call ADCBuffer2AudioBuffer(1)
                       Call CompareDataTrace(1)
                           
                       rv3 = WMACompareResult
                 Else
                   rv3 = 4
                 End If
                 Call LabelMenu(3, rv3, rv2)
                  
                 '========================================================
                 '
                 '  Recode Raw data to pattern file
                 '
                 '========================================================
                  
                  
                 If Tester.AudioRecordCheck.Value = 1 Then
                                       
                     Open App.Path & "\MP3Pattern\MP3.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, MP3RawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, MP3RawData(1, i)
                         Next i
                      
    
                    Close #4
                
                
                                  
                     Open App.Path & "\WMAPattern\WMA.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, WMARawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, WMARawData(1, i)
                         Next i
                      
    
                    Close #4
                    
                 End If
                 
                 If rv2 * rv3 = 1 Then
                 
                   ' Tester.Cls
                    Tester.Print "MP3Matchmode="; MP3MatchIndex
                    Tester.Print "WMAMatchmode="; WMAMatchIndex
                 End If

                 dma2205_4.Show
                 dma2205_4.ShowData 0, 6000
                
                 Card2214Exist = 0
                  D2K_Release_Card (Card2214)
                  CardResult = DO_WritePort(card, Channel_P1A, &HC0)  ' Allen debug
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
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
                ElseIf rv0 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub


Public Sub LoadAU3150JLF27AudioPattern()



'==========================================
'  MP3
'===========================================
Call MP3LoadDataSub("\MP31", 1, 1)

Call MP3LoadDataSub("\MP32", 2, 1)

Call MP3LoadDataSub("\MP33", 3, 1)

Call MP3LoadDataSub("\MP34", 4, 1)

 Call MP3LoadDataSub("\MP35", 5, 1)

 Call MP3LoadDataSub("\MP36", 6, 1)
 Call MP3LoadDataSub("\MP37", 7, 1)

Call MP3LoadDataSub("\MP38", 8, 1)

Call MP3LoadDataSub("\MP39", 9, 1)

Call MP3LoadDataSub("\MP3A", 10, 1)

' Call MP3LoadDataSub("\MP3B", 11, 1)

 ' Call MP3LoadDataSub("\MP3C", 12, 1)

' Call MP3LoadDataSub("\MP3D", 13, 1)

'  Call MP3LoadDataSub("\MP3E", 14, 1)

'  Call MP3LoadDataSub("\MP3F", 15, 1)

'  Call MP3LoadDataSub("\MP3G", 16, 1)

' Call MP3LoadDataSub("\MP3H", 17, 1)

' Call MP3LoadDataSub("\MP3I", 18, 1)

' Call MP3LoadDataSub("\MP3J", 19, 1)

' Call MP3LoadDataSub("\MP3K", 20, 1)







'==========================================
'  MP3
'===========================================
Call WMALoadDataSub("\WMA1", 1, 1)

Call WMALoadDataSub("\WMA2", 2, 1)

Call WMALoadDataSub("\WMA3", 3, 1)

Call WMALoadDataSub("\WMA4", 4, 1)

Call WMALoadDataSub("\WMA5", 5, 1)

Call WMALoadDataSub("\WMA6", 6, 1)
Call WMALoadDataSub("\WMA7", 7, 1)

Call WMALoadDataSub("\WMA8", 8, 1)

Call WMALoadDataSub("\WMA9", 9, 1)

Call WMALoadDataSub("\WMAA", 10, 1)

'Call WMALoadDataSub("\WMAB", 11, 1)

 'Call WMALoadDataSub("\WMAC", 12, 1)

 'Call WMALoadDataSub("\WMAD", 13, 1)

 'Call WMALoadDataSub("\WMAE", 14, 1)

 'Call WMALoadDataSub("\WMAF", 15, 1)

 'Call WMALoadDataSub("\WMAG", 16, 1)

 'Call WMALoadDataSub("\WMAH", 17, 1)

 'Call WMALoadDataSub("\WMAI", 18, 1)

 'Call WMALoadDataSub("\WMAJ", 19, 1)

 'Call WMALoadDataSub("\WMAK", 20, 1)


End Sub

Public Sub LoadAU3152ALF24AudioPattern()



'==========================================
'  MP3
'===========================================
Call MP3LoadDataSub("\MP31", 1, 1)

Call MP3LoadDataSub("\MP32", 2, 1)

Call MP3LoadDataSub("\MP33", 3, 1)

Call MP3LoadDataSub("\MP34", 4, 1)

Call MP3LoadDataSub("\MP35", 5, 1)

Call MP3LoadDataSub("\MP36", 6, 1)
Call MP3LoadDataSub("\MP37", 7, 1)

Call MP3LoadDataSub("\MP38", 8, 1)

Call MP3LoadDataSub("\MP39", 9, 1)

Call MP3LoadDataSub("\MP3A", 10, 1)
'==========================================
'  MP3
'===========================================
Call WMALoadDataSub("\WMA1", 1, 1)

Call WMALoadDataSub("\WMA2", 2, 1)

Call WMALoadDataSub("\WMA3", 3, 1)

Call WMALoadDataSub("\WMA4", 4, 1)

Call WMALoadDataSub("\WMA5", 5, 1)

Call WMALoadDataSub("\WMA6", 6, 1)

Call WMALoadDataSub("\WMA7", 7, 1)

Call WMALoadDataSub("\WMA8", 8, 1)

Call WMALoadDataSub("\WMA9", 9, 1)

Call WMALoadDataSub("\WMAA", 10, 1)

Call WMALoadDataSub("\WMAB", 11, 1)

End Sub


Public Sub LoadAU3152CLF24AudioPattern()
'==========================================
'  MP3
'===========================================
Call MP3LoadDataSub("\MP31", 1, 1)

Call MP3LoadDataSub("\MP32", 2, 1)

Call MP3LoadDataSub("\MP33", 3, 1)

Call MP3LoadDataSub("\MP34", 4, 1)

Call MP3LoadDataSub("\MP35", 5, 1)

Call MP3LoadDataSub("\MP36", 6, 1)
Call MP3LoadDataSub("\MP37", 7, 1)

Call MP3LoadDataSub("\MP38", 8, 1)

Call MP3LoadDataSub("\MP39", 9, 1)

Call MP3LoadDataSub("\MP3A", 10, 1)

Call MP3LoadDataSub("\MP3B", 11, 1)
'==========================================
'  MP3
'===========================================
Call WMALoadDataSub("\WMA1", 1, 1)

Call WMALoadDataSub("\WMA2", 2, 1)

Call WMALoadDataSub("\WMA3", 3, 1)

Call WMALoadDataSub("\WMA4", 4, 1)

Call WMALoadDataSub("\WMA5", 5, 1)

Call WMALoadDataSub("\WMA6", 6, 1)

Call WMALoadDataSub("\WMA7", 7, 1)

Call WMALoadDataSub("\WMA8", 8, 1)

Call WMALoadDataSub("\WMA9", 9, 1)

Call WMALoadDataSub("\WMAA", 10, 1)

Call WMALoadDataSub("\WMAB", 11, 1)

End Sub

Public Sub AU3152CLF24_1_SlotTest_Noise()
Dim i As Integer
  
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                 '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &H80)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                 CardResult = DO_WritePort(card, Channel_P1A, &H2F)   ' usb Power Off
                 Call MsecDelay(0.8)    'power on time
                  
                 Call MsecDelay(2.5)     'power on time
                  
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
            
                ClosePipe
                LBA = LBA + 1
                OldLBa = LBA
                LBA = LBA + 20000 ' 20000 mp3 shift
                
                rv0 = CBWTest_New_31303(0, 1, "vid_058f")
                LBA = OldLBa
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                Tester.Cls
                Tester.Print "rv0="; rv0
                
                
                 
                If rv0 = 1 Then
                    
                          '=================================
                          '   begin load pattern
                          '=================================
                          
                         
                          If OldChipName <> ChipName Then
                                 
                                Call LoadAU3152CLF24AudioPattern
                                OldChipName = ChipName
                          End If
                                                      
              
                          
                          '==========================================
                          '  do a low-high-low pulse at play button
                          '==========================================
                         
                            CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                             
                             Call MsecDelay(0.5)
                             
                             CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                               
                             Call MsecDelay(4)    'power on time and DAC initial time
                               
                             Call MsecDelay(2.5)     'for 3150K power on time

                            CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                          '==========================================
                          '  Get ADC Data and compare result
                          '==========================================
                  
                           Call DACGetData2ChannelLowTriggle64
                           Call ADCBuffer2AudioBuffer(0)
                           Call CompareData(0)
                           
                           rv2 = MP3CompareResult
                  Else
                   rv2 = 4
                 End If
                 
                 Call LabelMenu(2, rv2, rv0)
                 

               
                If rv2 = 1 Or Tester.AudioRecordCheck.Value = 1 Then
                     
                    '=========================================
                    '
                    '   Start Next one song WMA
                    '
                    '========================================
                
                      CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                      Tester.Print "next one press"
                      Call MsecDelay(0.5)
                      CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                      
                      Tester.Print "next one up"
                
                         
                      '=========================================
                      '
                      '   End Next one song WMA
                      '
                      '========================================
                    
                 
                       Call DACGetData2ChannelLowTriggle64
                       Call ADCBuffer2AudioBuffer(1)
                       Call CompareData(1)
                           
                       rv3 = WMACompareResult
                 Else
                   rv3 = 4
                 End If
                 Call LabelMenu(3, rv3, rv2)
                  
                 '========================================================
                 '
                 '  Recode Raw data to pattern file
                 '
                 '========================================================
                  
                  
                 If Tester.AudioRecordCheck.Value = 1 Then
                                       
                     Open App.Path & "\MP3Pattern\MP3.txt" For Output As #4
    
                         For i = 0 To 99
                           Print #4, MP3RawData(0, i)
                         Next i
                          For i = 0 To 99
                           Print #4, MP3RawData(1, i)
                         Next i
                      
    
                    Close #4
                
                
                                  
                     Open App.Path & "\WMAPattern\WMA.txt" For Output As #4
    
                         For i = 0 To 99
                           Print #4, WMARawData(0, i)
                         Next i
                          For i = 0 To 99
                           Print #4, WMARawData(1, i)
                         Next i
                      
    
                    Close #4
                    
                 End If
                 
                 If rv2 * rv3 = 1 Then
                 
                   ' Tester.Cls
                    Tester.Print "MP3Matchmode="; MP3MatchIndex
                    Tester.Print "WMAMatchmode="; WMAMatchIndex
                 End If

                 dma2205_4.Show
                 dma2205_4.ShowData 0, 100
                
                 Card2214Exist = 0
                  D2K_Release_Card (Card2214)
                  
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
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
                ElseIf rv0 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub

Public Sub LoadAudioPattern()



'==========================================
'  MP3
'===========================================
Call MP3LoadDataSub("\MP31", 1, 1)

Call MP3LoadDataSub("\MP32", 2, 1)

Call MP3LoadDataSub("\MP33", 3, 1)

Call MP3LoadDataSub("\MP34", 4, 1)

Call MP3LoadDataSub("\MP35", 5, 1)
'==========================================
'  MP3
'===========================================
Call WMALoadDataSub("\WMA1", 1, 1)

Call WMALoadDataSub("\WMA2", 2, 1)

'Call WMALoadDataSub("\WMA3", 3, 1)

End Sub

Public Sub WMALoadDataSub(PatternName As String, WMAPatternIndexSub As Byte, FileNo As Byte)
On Error Resume Next
Dim i As Integer
Dim TmpData1 As Integer
Dim TmpData2 As Integer
Dim TmpData3 As Integer
Dim TmpData4 As Integer
Dim TmpData5 As Integer

If FileNo = 1 Then
 Open App.Path & "\WMAPattern\" & ChipName & PatternName & "1.txt" For Input As #4

                     For i = 0 To 5999
                       Input #4, WMAPattern(WMAPatternIndexSub - 1, 0, i)
                     Next i
                      For i = 0 To 5999
                       Input #4, WMAPattern(WMAPatternIndexSub - 1, 1, i)
                     Next i
                  
 Close #4
End If


If FileNo = 2 Then
Open App.Path & "\WMAPattern\" & ChipName & PatternName & "1.txt" For Input As #4
Open App.Path & "\WMAPattern\" & ChipName & PatternName & "2.txt" For Input As #5

                     For i = 0 To 5999
                       Input #4, TmpData1
                       Input #5, TmpData2
                       WMAPattern(WMAPatternIndexSub - 1, 0, i) = CInt((CSng(TmpData1) + CSng(TmpData2)) * 0.5)
                     Next i
                    For i = 0 To 5999
                      Input #4, TmpData1
                      Input #5, TmpData2
                      WMAPattern(WMAPatternIndexSub - 1, 1, i) = CInt((CSng(TmpData1) + CSng(TmpData2)) * 0.5)
                       
                     Next i
                  
 Close #4
 Close #5
End If

 
 
 WMAPatternIndex = WMAPatternIndexSub
End Sub

Public Sub MP3LoadDataSub(PatternName As String, MP3PatternIndexSub As Byte, FileNo As Byte)
On Error Resume Next
Dim i As Integer
Dim TmpData1 As Integer
Dim TmpData2 As Integer
Dim TmpData3 As Integer
Dim TmpData4 As Integer
Dim TmpData5 As Integer

If FileNo = 1 Then
 Open App.Path & "\MP3Pattern\" & ChipName & PatternName & "1.txt" For Input As #4

                     For i = 0 To 5999
                       Input #4, MP3Pattern(MP3PatternIndexSub - 1, 0, i)
                     Next i
                      For i = 0 To 5999
                       Input #4, MP3Pattern(MP3PatternIndexSub - 1, 1, i)
                     Next i
                  
 Close #4
End If


If FileNo = 2 Then
Open App.Path & "\MP3Pattern\" & ChipName & PatternName & "1.txt" For Input As #4
Open App.Path & "\MP3Pattern\" & ChipName & PatternName & "2.txt" For Input As #5

                     For i = 0 To 5999
                       Input #4, TmpData1
                       Input #5, TmpData2
                       MP3Pattern(MP3PatternIndexSub - 1, 0, i) = CInt((CSng(TmpData1) + CSng(TmpData2)) * 0.5)
                     Next i
                    For i = 0 To 5999
                      Input #4, TmpData1
                      Input #5, TmpData2
                      MP3Pattern(MP3PatternIndexSub - 1, 1, i) = CInt((CSng(TmpData1) + CSng(TmpData2)) * 0.5)
                       
                     Next i
                  
 Close #4
 Close #5
End If



If FileNo = 4 Then
Open App.Path & "\MP3Pattern\" & ChipName & PatternName & "1.txt" For Input As #4
Open App.Path & "\MP3Pattern\" & ChipName & PatternName & "2.txt" For Input As #5
Open App.Path & "\MP3Pattern\" & ChipName & PatternName & "3.txt" For Input As #6
Open App.Path & "\MP3Pattern\" & ChipName & PatternName & "4.txt" For Input As #7

                     For i = 0 To 5999
                       Input #4, TmpData1
                       Input #5, TmpData2
                       Input #4, TmpData3
                       Input #5, TmpData4
                       MP3Pattern(MP3PatternIndexSub - 1, 0, i) = CInt((CSng(TmpData1) + CSng(TmpData2) + CSng(TmpData3) + CSng(TmpData4)) * 0.25)
                     Next i
                    For i = 0 To 5999
                      Input #4, TmpData1
                      Input #5, TmpData2
                      Input #4, TmpData3
                      Input #5, TmpData4
                      MP3Pattern(MP3PatternIndexSub - 1, 1, i) = CInt((CSng(TmpData1) + CSng(TmpData2) + CSng(TmpData3) + CSng(TmpData4)) * 0.25)
                       
                     Next i
                  
 Close #4
 Close #5
  Close #6
 Close #7
End If

 
 
 MP3PatternIndex = MP3PatternIndexSub
End Sub
Public Sub AU3150K_1_SlotTest_Noise()
Dim i As Integer
  
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                 '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &H80)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                 CardResult = DO_WritePort(card, Channel_P1A, &H2F)   ' usb Power Off
                 Call MsecDelay(0.8)    'power on time
                  
                 Call MsecDelay(2.5)     'power on time
                  
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
            
                ClosePipe
                LBA = LBA + 1
                OldLBa = LBA
                LBA = LBA + 20000 ' 20000 mp3 shift
                
                rv0 = CBWTest_New_31303(0, 1, "vid_058f")
                LBA = OldLBa
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                Tester.Cls
                Tester.Print "rv0="; rv0
                
                
                 
                If rv0 = 1 Then
                    
                          '=================================
                          '   begin load pattern
                          '=================================
                          
                         
                          If OldChipName <> ChipName Then
                                 
                                Call LoadAudioPattern
                                OldChipName = ChipName
                          End If
                                                      
              
                          
                          '==========================================
                          '  do a low-high-low pulse at play button
                          '==========================================
                         
                            CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                             
                             Call MsecDelay(0.5)
                             
                             CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                               
                             Call MsecDelay(4)    'power on time and DAC initial time
                               
                             Call MsecDelay(2.5)     'for 3150K power on time

                            CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                          '==========================================
                          '  Get ADC Data and compare result
                          '==========================================
                  
                           Call DACGetData2Channel
                           Call ADCBuffer2AudioBuffer(0)
                           Call CompareData(0)
                           
                           rv2 = MP3CompareResult
                  Else
                   rv2 = 4
                 End If
                 
                 Call LabelMenu(2, rv2, rv0)
                 

               
                If rv2 = 1 Or Tester.AudioRecordCheck.Value = 1 Then
                     
                    '=========================================
                    '
                    '   Start Next one song WMA
                    '
                    '========================================
                
                      CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                      Tester.Print "next one press"
                      Call MsecDelay(0.5)
                      CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                      
                      Tester.Print "next one up"
                
         
                      '=========================================
                      '
                      '   End Next one song WMA
                      '
                      '========================================
                    
                 
                       Call DACGetData2Channel
                       Call ADCBuffer2AudioBuffer(1)
                       Call CompareData(1)
                           
                       rv3 = WMACompareResult
                 Else
                   rv3 = 4
                 End If
                 Call LabelMenu(3, rv3, rv2)
                  
                 '========================================================
                 '
                 '  Recode Raw data to pattern file
                 '
                 '========================================================
                  
                  
                 If Tester.AudioRecordCheck.Value = 1 Then
                                       
                     Open App.Path & "\MP3Pattern\MP3.txt" For Output As #4
    
                         For i = 0 To 99
                           Print #4, MP3RawData(0, i)
                         Next i
                          For i = 0 To 99
                           Print #4, MP3RawData(1, i)
                         Next i
                      
    
                    Close #4
                
                
                                  
                     Open App.Path & "\WMAPattern\WMA.txt" For Output As #4
    
                         For i = 0 To 99
                           Print #4, WMARawData(0, i)
                         Next i
                          For i = 0 To 99
                           Print #4, WMARawData(1, i)
                         Next i
                      
    
                    Close #4
                    
                 End If
                 
                 If rv2 * rv3 = 1 Then
                 
                   ' Tester.Cls
                    Tester.Print "MP3Matchmode="; MP3MatchIndex
                    Tester.Print "WMAMatchmode="; WMAMatchIndex
                 End If

                 dma2205_4.Show
                 dma2205_4.ShowData 0, 100
                
                 Card2214Exist = 0
                  D2K_Release_Card (Card2214)
                  
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
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
                ElseIf rv0 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub


Public Sub AU3150ALF24_2_SlotTest_Noise()
Dim i As Integer
  
                
                
                
                 If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                 End If
                
                
                 '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &H80)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                 CardResult = DO_WritePort(card, Channel_P1A, &H2F)   ' usb Power Off
                
                 Call MsecDelay(0.8)    'power on time
                 
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
                 
                
                
                ClosePipe
                LBA = LBA + 1
                OldLBa = LBA
                 LBA = LBA + 20000 ' 20000 mp3 shift
                rv0 = CBWTest_New_3130(0, 1, "vid_058f")
               
                 LBA = OldLBa
               
                 Call LabelMenu(0, rv0, 1)
                ClosePipe
                 rv1 = CBWTest_New_31302(1, rv0, "vid_058f")
                 Call LabelMenu(1, rv1, rv0)
                 ClosePipe
                 
                 Tester.Print " Flash R/W test rv0="; rv0
                 
                  Tester.Print "SD R/W teswt rv1="; rv1
                
                
                
                
                 
                If rv0 * rv1 = 1 Then
                    
                          '=================================
                          '   begin load pattern
                          '=================================
                          
                         
                          If OldChipName <> ChipName Then
                                 
                                Call LoadAudioPattern
                                OldChipName = ChipName
                          End If
                                                      
              
                          
                          '==========================================
                          '  do a low-high-low pulse at play button
                          '==========================================
                         
                            CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                             
                             Call MsecDelay(0.5)
                             
                             CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                               
                             Call MsecDelay(4)    'power on time and DAC initial time
                               
                             Call MsecDelay(2.5)     'for 3150K power on time

                            CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                          '==========================================
                          '  Get ADC Data and compare result
                          '==========================================
                  
                           Call DACGetData2ChannelLowTriggle
                           Call ADCBuffer2AudioBuffer(0)
                           Call CompareData(0)
                           
                           rv2 = MP3CompareResult
                  Else
                   rv2 = 4
                 End If
                 
                 Call LabelMenu(2, rv2, rv0)
                 

               
                If rv2 = 1 Or Tester.AudioRecordCheck.Value = 1 Then
                     
                    '=========================================
                    '
                    '   Start Next one song WMA
                    '
                    '========================================
                
                      CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                      Tester.Print "next one press"
                      Call MsecDelay(0.5)
                      CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                      
                      Tester.Print "next one up"
                
         
                      '=========================================
                      '
                      '   End Next one song WMA
                      '
                      '========================================
                    
                 
                       Call DACGetData2ChannelLowTriggle
                       Call ADCBuffer2AudioBuffer(1)
                       Call CompareData(1)
                           
                       rv3 = WMACompareResult
                 Else
                   rv3 = 4
                 End If
                 Call LabelMenu(3, rv3, rv2)
                  
                 '========================================================
                 '
                 '  Recode Raw data to pattern file
                 '
                 '========================================================
                  
                  
                 If Tester.AudioRecordCheck.Value = 1 Then
                                       
                     Open App.Path & "\MP3Pattern\MP3.txt" For Output As #4
    
                         For i = 0 To 99
                           Print #4, MP3RawData(0, i)
                         Next i
                          For i = 0 To 99
                           Print #4, MP3RawData(1, i)
                         Next i
                      
    
                    Close #4
                
                
                                  
                     Open App.Path & "\WMAPattern\WMA.txt" For Output As #4
    
                         For i = 0 To 99
                           Print #4, WMARawData(0, i)
                         Next i
                          For i = 0 To 99
                           Print #4, WMARawData(1, i)
                         Next i
                      
    
                    Close #4
                    
                 End If
                 
                 If rv2 * rv3 = 1 Then
                 
                   ' Tester.Cls
                    Tester.Print "MP3Matchmode="; MP3MatchIndex
                    Tester.Print "WMAMatchmode="; WMAMatchIndex
                 End If

                 dma2205_4.Show
                 dma2205_4.ShowData 0, 100
                
                 Card2214Exist = 0
                  D2K_Release_Card (Card2214)
                  
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
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
                ElseIf rv0 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub


Public Function DACGetData2ChannelUi() As Integer
'On Error Resume Next
  Dim result As Long
  Dim Status As Byte
  Dim i As Long, k As Long
  Dim ReTrgCnt As Integer, MCnt As Integer, PostCnt As Integer
  Dim TimerSrc As Integer
  Dim cobTrigModeListIndex As Integer
  Dim cobTrigSrcListIndex As Integer
  Dim cobTrigPolListIndex As Integer
  Dim cobClkSrcListIndex As Integer
  Dim ErrCount As Integer
  Dim OldCount As Integer
  Dim CoutinueCount As Integer
  Dim cobTrigLevelText As String
  Dim cobTrigL_LevelText As String
  Dim gnClkDiv As Integer
  Dim Id As Integer
  Dim startPos As Long
  Dim OldTimer
  Dim ADCPostCnt As Integer
  Dim ADCdma_size As Long
  
   ' dma2205.Hide
  ADCPostCnt = 1 ' post trigger 500
  ADCdma_size = 100 ' scan count 2000-- 80000
  ch_cnt = 2
  ADCch_num(0) = 0  ' channel 0
  
  ADCch_range(0) = 18  ' BiPolar 2.5V, 1 base
  ADCch_ref(0) = 0    ' Differentail
  
  
  ADCch_num(1) = 1  ' channel 0
  
  
   ADCch_range(1) = 18  ' BiPolar 2.5V, 1 base
  ADCch_ref(0) = 0    ' Differentail
  
  '=================================
  '  DMA setting
  '=================================
  
        cobTrigModeListIndex = 0
        cobTrigSrcListIndex = 1
        cobTrigPolListIndex = 1
        cobClkSrcListIndex = 0
        cobTrigPolListIndex = 1
        cobTrigLevelText = "200"
        cobTrigL_LevelText = "128"
        gnClkDiv = 5000
   
       
    Call MsecDelay(0.5)   'for AU3150 unstable chip
    

   If Card2214Exist = 0 Then
        Card2214 = D2K_Register_Card(DAQ_2214, 0)
        
          For i = 0 To 1 Step 1
                result = D2K_AI_CH_Config(Card2214, ADCch_num(i), ADCch_range(i) Or (ADCch_ref(i) * 256))
          Next
          
   
          
         
    End If
   
   
  For i = 0 To 99
  ADCBuffer(i) = 0
  Next i
  
   
        result = D2K_AI_Config(Card2214, 0, (CLng(cobTrigModeListIndex) * 8) Or CLng(cobTrigSrcListIndex) Or (CLng(cobTrigPolListIndex) * 4096), ADCPostCnt, 0, 1, 1)
          If result <> 0 Then
            MsgBox "2214 configure fail 1"
            Exit Function
          End If
   
   
  
   result = D2K_AI_ContBufferSetup(Card2214, ADCBuffer(0), ADCdma_size * ch_cnt, Id)
'   Tester.Print "result2 ="; result; "Id="; Id
   
    result = D2K_AIO_Config(Card2214, cobClkSrcListIndex, (CLng(cobTrigPolListIndex) * 256) Or CH0ATRIG, CLng(cobTrigLevelText), CLng(cobTrigL_LevelText))
            If result <> 0 Then
             MsgBox "2214 configure fail 2"
             Exit Function
            End If
  'Tester.Print "result4="; result
  result = D2K_AI_ContReadMultiChannels(Card2214, ch_cnt, ADCch_num(0), Id, ADCdma_size, gnClkDiv * ch_cnt, gnClkDiv, ASYNCH_OP)
     
 '  Tester.Print "read AD begin ="; result
   Card2214Exist = 1
  
  
  
  
  Status = 0
  bStop = 0
  OldTimer = Timer
  While Status = 0 And bStop = 0 And (Timer - OldTimer < 10)
     DoEvents
    result = D2K_AI_AsyncCheck(Card2214, Status, count1)
    
  Wend
  result = D2K_AI_AsyncClear(Card2214, startPos, count1)
  Tester.Print "result5 ="; result;
   Tester.Print "read AD end "
  
 End Function
 
 Public Function DACGetData1ChannelUi() As Integer
'On Error Resume Next
  Dim result As Long
  Dim Status As Byte
  Dim i As Long, k As Long
  Dim ReTrgCnt As Integer, MCnt As Integer, PostCnt As Integer
  Dim TimerSrc As Integer
  Dim cobTrigModeListIndex As Integer
  Dim cobTrigSrcListIndex As Integer
  Dim cobTrigPolListIndex As Integer
  Dim cobClkSrcListIndex As Integer
  Dim ErrCount As Integer
  Dim OldCount As Integer
  Dim CoutinueCount As Integer
  Dim cobTrigLevelText As String
  Dim cobTrigL_LevelText As String
  Dim gnClkDiv As Integer
  Dim Id As Integer
  Dim startPos As Long
  Dim OldTimer
  Dim ADCPostCnt As Integer
  Dim ADCdma_size As Long
  
   ' dma2205.Hide
  ADCPostCnt = 1 ' post trigger 500
  ADCdma_size = 1000 ' scan count 2000-- 80000,AD_B_5_V = 2,AD_U_5_V = 16
  ch_cnt = 1
  ADCch_num(0) = 0  ' channel 0
  
  ADCch_range(0) = 2  ' BiPolar 5V, 1 base
  ADCch_ref(0) = 0    ' Differentail
  
  
  ADCch_num(1) = 1  ' channel 0
  
  
   ADCch_range(1) = 18  ' BiPolar 2.5V, 1 base
  ADCch_ref(0) = 0    ' Differentail
  
  '=================================
  '  DMA setting
  '=================================
  
        cobTrigModeListIndex = 0
        cobTrigSrcListIndex = 1
        cobTrigPolListIndex = 1
        cobClkSrcListIndex = 0
        cobTrigPolListIndex = 1
        cobTrigLevelText = "200"
        cobTrigL_LevelText = "128"
        gnClkDiv = 5000
   
       
    Call MsecDelay(0.5)   'for AU3150 unstable chip
    

   If Card2214Exist = 0 Then
        Card2214 = D2K_Register_Card(DAQ_2214, 0)
        
           For i = 0 To 0 Step 1
                result = D2K_AI_CH_Config(Card2214, ADCch_num(i), ADCch_range(i) Or (ADCch_ref(i) * 256))
           Next
          
   
          
         
    End If
   
   
  For i = 0 To 999
  ADCBuffer(i) = 0
  Next i
  
   
        result = D2K_AI_Config(Card2214, 0, (CLng(cobTrigModeListIndex) * 8) Or CLng(cobTrigSrcListIndex) Or (CLng(cobTrigPolListIndex) * 4096), ADCPostCnt, 0, 1, 1)
          If result <> 0 Then
            MsgBox "2214 configure fail 1"
            Exit Function
          End If
   
   
  
   result = D2K_AI_ContBufferSetup(Card2214, ADCBuffer(0), ADCdma_size * ch_cnt, Id)
'   Tester.Print "result2 ="; result; "Id="; Id
   
    result = D2K_AIO_Config(Card2214, cobClkSrcListIndex, (CLng(cobTrigPolListIndex) * 256) Or CH0ATRIG, CLng(cobTrigLevelText), CLng(cobTrigL_LevelText))
            If result <> 0 Then
             MsgBox "2214 configure fail 2"
             Exit Function
            End If
  'Tester.Print "result4="; result
  result = D2K_AI_ContReadMultiChannels(Card2214, ch_cnt, ADCch_num(0), Id, ADCdma_size, gnClkDiv * ch_cnt, gnClkDiv, ASYNCH_OP)
     
 '  Tester.Print "read AD begin ="; result
   Card2214Exist = 1
  
  
  
  
  Status = 0
  bStop = 0
  OldTimer = Timer
  While Status = 0 And bStop = 0 And (Timer - OldTimer < 10)
     DoEvents
    result = D2K_AI_AsyncCheck(Card2214, Status, count1)
    
  Wend
  result = D2K_AI_AsyncClear(Card2214, startPos, count1)
  Tester.Print "result5 ="; result;
   Tester.Print "read AD end "
  
 End Function


Public Sub AU3150KLF27_1_SlotTest_Noise()
Dim i As Integer
  
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                 '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &H80)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                 CardResult = DO_WritePort(card, Channel_P1A, &H2F)   ' usb Power Off
                 Call MsecDelay(0.8)    'power on time
                  
                 Call MsecDelay(2.5)     'power on time
                  
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
            
                ClosePipe
                LBA = LBA + 1
                OldLBa = LBA
                LBA = LBA + 40000 ' 20000 mp3 shift
                
                rv0 = CBWTest_New_31303(0, 1, "vid_058f")
                LBA = OldLBa
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                Tester.Cls
                Tester.Print "rv0="; rv0
                 
                 
                If rv0 = 1 Then
                    
                          '=================================
                          '   begin load pattern
                          '=================================
                          
                         
                          If OldChipName <> ChipName Then
                                 
                                Call LoadAU3150JLF27AudioPattern
                                OldChipName = ChipName
                          End If
                                                      
              
                          
                          '==========================================
                          '  do a low-high-low pulse at play button
                          '==========================================
                         
                            CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                             
                             Call MsecDelay(0.5)
                             
                             CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                               
                              Call MsecDelay(4)    'power on time and DAC initial time
                               
                             Call MsecDelay(2.5)     'for 3150K power on time

                            CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                          '==========================================
                          '  Get ADC Data and compare result
                          '==========================================
                   Call MsecDelay(1.5)
                           Call DACGetData2ChannelUpTriggle
                           Call ADCBuffer2AudioBuffer(0)
                           Call CompareDataTrace(0)
                           
                           rv2 = MP3CompareResult
                  Else
                   rv2 = 4
                 End If
                 
                 Call LabelMenu(2, rv2, rv0)
                 

               
                If (rv0 = 1) And (rv2 = 1 Or Tester.AudioRecordCheck.Value = 1) Then
                  
                      D2K_Release_Card (Card2214)
                     Card2214Exist = 0
                     
                    '=========================================
                    '
                    '   Start Next one song WMA
                    '
                    '========================================
                     Tester.Print "next one press"
                      CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                    
                      Call MsecDelay(0.2)
                      CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                  
                      Tester.Print "next one up"
                        
                        
                         
                      '=========================================
                      '
                      '   End Next one song WMA
                      '
                      '========================================
                    
                     ' Call MsecDelay(1#)
                     '  CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                       Call DACGetData2ChannelUpTriggle
                       Call ADCBuffer2AudioBuffer(1)
                       Call CompareDataTrace(1)
                           
                       rv3 = WMACompareResult
                 Else
                   rv3 = 4
                 End If
                 Call LabelMenu(3, rv3, rv2)
                  
                 '========================================================
                 '
                 '  Recode Raw data to pattern file
                 '
                 '========================================================
                  
                  
                 If Tester.AudioRecordCheck.Value = 1 Then
                                       
                     Open App.Path & "\MP3Pattern\MP3.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, MP3RawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, MP3RawData(1, i)
                         Next i
                      
    
                    Close #4
                
                
                                  
                     Open App.Path & "\WMAPattern\WMA.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, WMARawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, WMARawData(1, i)
                         Next i
                      
    
                    Close #4
                    
                 End If
                 
                 If rv2 * rv3 = 1 Then
                 
                   ' Tester.Cls
                    Tester.Print "MP3Matchmode="; MP3MatchIndex
                    Tester.Print "WMAMatchmode="; WMAMatchIndex
                 End If

                 dma2205_4.Show
                 dma2205_4.ShowData 0, 6000
                
                 Card2214Exist = 0
                  D2K_Release_Card (Card2214)
                  
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
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
                ElseIf rv0 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub
Public Sub AU3150JLF28_1_SlotTest_Noise()
Dim i As Integer
  
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                 '=========================================
                '    POWER on
                '=========================================
               '   CardResult = DO_WritePort(card, Channel_P1A, &H80)
                    CardResult = DO_WritePort(card, Channel_P1A, &HC0)  ' Allen debug
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                CardResult = DO_WritePort(card, Channel_P1A, &H6F)
              '    CardResult = DO_WritePort(card, Channel_P1A, &H2F)   ' usb Power Off
                  Call MsecDelay(0.8)    'power on time
                  
              '   Call MsecDelay(1#)      'power on time
                  
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
            
                ClosePipe
                LBA = LBA + 1
                OldLBa = LBA
       
                
                LBA = LBA + 40000 ' 20000 mp3 shift
                
                rv0 = CBWTest_New_31303(0, 1, "vid_058f")
                LBA = OldLBa
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 Tester.Cls
                Tester.Print "LBA="; LBA
                Tester.Print "rv0="; rv0
                 
                If rv0 = 1 Then
                    
                          '=================================
                          '   begin load pattern
                          '=================================
                          
                         
                          If OldChipName <> ChipName Then
                                 
                                Call LoadAU3150JLF27AudioPattern
                                OldChipName = ChipName
                          End If
                                                      
              
                          
                          '==========================================
                          '  do a low-high-low pulse at play button
                          '==========================================
                         
                             
                             
                             CardResult = DO_WritePort(card, Channel_P1A, &HCF)
                             
                             Call MsecDelay(0.3)
                             CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                                 Call MsecDelay(2.5)
                                'Call MsecDelay(4.5)
                              CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                            '   CardResult = DO_WritePort(card, Channel_P1A, &HDF)
                          '    Call MsecDelay(4)    'power on time and DAC initial time
                               
                            ' Call MsecDelay(2.5)     'for 3150K power on time
                               Call MsecDelay(0.3)
                             CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                            
                            ' CardResult = DO_WritePort(card, Channel_P1A, &HCF)
                          '==========================================
                          '  Get ADC Data and compare result
                          '==========================================
                   Call MsecDelay(1.5)
                           Call DACGetData2ChannelUpTriggle
                           Call ADCBuffer2AudioBuffer(0)
                           Call CompareDataTrace(0)
                           
                           rv2 = MP3CompareResult
                  Else
                   rv2 = 4
                 End If
                 
                 Call LabelMenu(2, rv2, rv0)
                 

               
                If (rv0 = 1) And (rv2 = 1 Or Tester.AudioRecordCheck.Value = 1) Then
                  
                      D2K_Release_Card (Card2214)
                     Card2214Exist = 0
                     
                    '=========================================
                    '
                    '   Start Next one song WMA
                    '
                    '========================================
                     Tester.Print "next one press"
                      CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                    
                      Call MsecDelay(0.2)
                      CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                  
                      Tester.Print "next one up"
                        
                        
                         
                      '=========================================
                      '
                      '   End Next one song WMA
                      '
                      '========================================
                    
                     ' Call MsecDelay(1#)
                     '  CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                       Call DACGetData2ChannelUpTriggle
                       Call ADCBuffer2AudioBuffer(1)
                       Call CompareDataTrace(1)
                           
                       rv3 = WMACompareResult
                 Else
                   rv3 = 4
                 End If
                 Call LabelMenu(3, rv3, rv2)
                  
                 '========================================================
                 '
                 '  Recode Raw data to pattern file
                 '
                 '========================================================
                  
                  
                 If Tester.AudioRecordCheck.Value = 1 Then
                                       
                     Open App.Path & "\MP3Pattern\MP3.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, MP3RawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, MP3RawData(1, i)
                         Next i
                      
    
                    Close #4
                
                
                                  
                     Open App.Path & "\WMAPattern\WMA.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, WMARawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, WMARawData(1, i)
                         Next i
                      
    
                    Close #4
                    
                 End If
                 
                 If rv2 * rv3 = 1 Then
                 
                   ' Tester.Cls
                    Tester.Print "MP3Matchmode="; MP3MatchIndex
                    Tester.Print "WMAMatchmode="; WMAMatchIndex
                 End If

                 dma2205_4.Show
                 dma2205_4.ShowData 0, 6000
                
                 Card2214Exist = 0
                  D2K_Release_Card (Card2214)
                  CardResult = DO_WritePort(card, Channel_P1A, &HC0)  ' Allen debug
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
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
                ElseIf rv0 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub
Public Sub AU3152ALF27_1_SlotTest_Noise()
Dim i As Integer
  
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                 '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &H80)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                 CardResult = DO_WritePort(card, Channel_P1A, &H2F)   ' usb Power Off
                 Call MsecDelay(0.8)    'power on time
                  
                 Call MsecDelay(2.5)     'power on time
                  
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
            
                ClosePipe
                LBA = LBA + 1
                OldLBa = LBA
                LBA = LBA + 40000 ' 20000 mp3 shift
                
                rv0 = CBWTest_New_31303(0, 1, "vid_058f")
                LBA = OldLBa
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                Tester.Cls
                Tester.Print "rv0="; rv0
                 
                 
                If rv0 = 1 Then
                    
                          '=================================
                          '   begin load pattern
                          '=================================
                          
                         
                          If OldChipName <> ChipName Then
                                 
                                Call LoadAU3150JLF27AudioPattern
                                OldChipName = ChipName
                          End If
                                                      
              
                          
                          '==========================================
                          '  do a low-high-low pulse at play button
                          '==========================================
                         
                            CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                             
                             Call MsecDelay(0.5)
                             
                             CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                               
                              Call MsecDelay(4)    'power on time and DAC initial time
                               
                             Call MsecDelay(2.5)     'for 3150K power on time

                            CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                          '==========================================
                          '  Get ADC Data and compare result
                          '==========================================
                   Call MsecDelay(1.5)
                           Call DACGetData2ChannelUpTriggle
                           Call ADCBuffer2AudioBuffer(0)
                           Call CompareDataTrace(0)
                           
                           rv2 = MP3CompareResult
                  Else
                   rv2 = 4
                 End If
                 
                 Call LabelMenu(2, rv2, rv0)
                 

               
                If (rv0 = 1) And (rv2 = 1 Or Tester.AudioRecordCheck.Value = 1) Then
                  
                      D2K_Release_Card (Card2214)
                     Card2214Exist = 0
                     
                    '=========================================
                    '
                    '   Start Next one song WMA
                    '
                    '========================================
                     Tester.Print "next one press"
                      CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                    
                      Call MsecDelay(0.2)
                      CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                  
                      Tester.Print "next one up"
                        
                        
                         
                      '=========================================
                      '
                      '   End Next one song WMA
                      '
                      '========================================
                    
                     ' Call MsecDelay(1#)
                     '  CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                       Call DACGetData2ChannelUpTriggle
                       Call ADCBuffer2AudioBuffer(1)
                       Call CompareDataTrace(1)
                           
                       rv3 = WMACompareResult
                 Else
                   rv3 = 4
                 End If
                 Call LabelMenu(3, rv3, rv2)
                  
                 '========================================================
                 '
                 '  Recode Raw data to pattern file
                 '
                 '========================================================
                  
                  
                 If Tester.AudioRecordCheck.Value = 1 Then
                                       
                     Open App.Path & "\MP3Pattern\MP3.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, MP3RawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, MP3RawData(1, i)
                         Next i
                      
    
                    Close #4
                
                
                                  
                     Open App.Path & "\WMAPattern\WMA.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, WMARawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, WMARawData(1, i)
                         Next i
                      
    
                    Close #4
                    
                 End If
                 
                 If rv2 * rv3 = 1 Then
                 
                   ' Tester.Cls
                    Tester.Print "MP3Matchmode="; MP3MatchIndex
                    Tester.Print "WMAMatchmode="; WMAMatchIndex
                 End If

                 dma2205_4.Show
                 dma2205_4.ShowData 0, 6000
                
                 Card2214Exist = 0
                  D2K_Release_Card (Card2214)
                  
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
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
                ElseIf rv0 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub

Public Sub AU3150LLF27_1_SlotTest_Noise()
Dim i As Integer
  
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                 '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &H80)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                 CardResult = DO_WritePort(card, Channel_P1A, &H2F)   ' usb Power Off
                 Call MsecDelay(0.8)    'power on time
                  
                 Call MsecDelay(2.5)     'power on time
                  
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
            
                ClosePipe
                LBA = LBA + 1
                OldLBa = LBA
                LBA = LBA + 40000 ' 20000 mp3 shift
                
                rv0 = CBWTest_New_31303(0, 1, "vid_058f")
                LBA = OldLBa
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                Tester.Cls
                Tester.Print "rv0="; rv0
                 
                 
                If rv0 = 1 Then
                    
                          '=================================
                          '   begin load pattern
                          '=================================
                          
                         
                          If OldChipName <> ChipName Then
                                 
                                Call LoadAU3150JLF27AudioPattern
                                OldChipName = ChipName
                          End If
                                                      
              
                          
                          '==========================================
                          '  do a low-high-low pulse at play button
                          '==========================================
                         
                            CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                             
                             Call MsecDelay(0.5)
                             
                             CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                               
                              Call MsecDelay(4)    'power on time and DAC initial time
                               
                             Call MsecDelay(2.5)     'for 3150K power on time

                            CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                          '==========================================
                          '  Get ADC Data and compare result
                          '==========================================
                   Call MsecDelay(1.5)
                           Call DACGetData2ChannelUpTriggle
                           Call ADCBuffer2AudioBuffer(0)
                           Call CompareDataTrace(0)
                           
                           rv2 = MP3CompareResult
                  Else
                   rv2 = 4
                 End If
                 
                 Call LabelMenu(2, rv2, rv0)
                 

               
                If (rv0 = 1) And (rv2 = 1 Or Tester.AudioRecordCheck.Value = 1) Then
                  
                      D2K_Release_Card (Card2214)
                     Card2214Exist = 0
                     
                    '=========================================
                    '
                    '   Start Next one song WMA
                    '
                    '========================================
                     Tester.Print "next one press"
                      CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                    
                      Call MsecDelay(0.2)
                      CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                  
                      Tester.Print "next one up"
                        
                        
                         
                      '=========================================
                      '
                      '   End Next one song WMA
                      '
                      '========================================
                    
                     ' Call MsecDelay(1#)
                     '  CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                       Call DACGetData2ChannelUpTriggle
                       Call ADCBuffer2AudioBuffer(1)
                       Call CompareDataTrace(1)
                           
                       rv3 = WMACompareResult
                 Else
                   rv3 = 4
                 End If
                 Call LabelMenu(3, rv3, rv2)
                  
                 '========================================================
                 '
                 '  Recode Raw data to pattern file
                 '
                 '========================================================
                  
                  
                 If Tester.AudioRecordCheck.Value = 1 Then
                                       
                     Open App.Path & "\MP3Pattern\MP3.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, MP3RawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, MP3RawData(1, i)
                         Next i
                      
    
                    Close #4
                
                
                                  
                     Open App.Path & "\WMAPattern\WMA.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, WMARawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, WMARawData(1, i)
                         Next i
                      
    
                    Close #4
                    
                 End If
                 
                 If rv2 * rv3 = 1 Then
                 
                   ' Tester.Cls
                    Tester.Print "MP3Matchmode="; MP3MatchIndex
                    Tester.Print "WMAMatchmode="; WMAMatchIndex
                 End If

                 dma2205_4.Show
                 dma2205_4.ShowData 0, 6000
                
                 Card2214Exist = 0
                  D2K_Release_Card (Card2214)
                  
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
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
                ElseIf rv0 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub


Public Sub AU3152CLF27_1_SlotTest_Noise()
Dim i As Integer
  
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                 '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &H80)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                 CardResult = DO_WritePort(card, Channel_P1A, &H2F)   ' usb Power Off
                 Call MsecDelay(0.8)    'power on time
                  
                 Call MsecDelay(2.5)     'power on time
                  
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
            
                ClosePipe
                LBA = LBA + 1
                OldLBa = LBA
                LBA = LBA + 40000 ' 20000 mp3 shift
                
                rv0 = CBWTest_New_31303(0, 1, "vid_058f")
                LBA = OldLBa
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                Tester.Cls
                Tester.Print "rv0="; rv0
                 
                 
                If rv0 = 1 Then
                    
                          '=================================
                          '   begin load pattern
                          '=================================
                          
                         
                          If OldChipName <> ChipName Then
                                 
                                Call LoadAU3150JLF27AudioPattern
                                OldChipName = ChipName
                          End If
                                                      
              
                          
                          '==========================================
                          '  do a low-high-low pulse at play button
                          '==========================================
                         
                            CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                             
                             Call MsecDelay(0.5)
                             
                             CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                               
                              Call MsecDelay(4)    'power on time and DAC initial time
                               
                             Call MsecDelay(2.5)     'for 3150K power on time

                            CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                          '==========================================
                          '  Get ADC Data and compare result
                          '==========================================
                   Call MsecDelay(1.5)
                           Call DACGetData2ChannelUpTriggle
                           Call ADCBuffer2AudioBuffer(0)
                           Call CompareDataTrace(0)
                           
                           rv2 = MP3CompareResult
                  Else
                   rv2 = 4
                 End If
                 
                 Call LabelMenu(2, rv2, rv0)
                 

               
                If (rv0 = 1) And (rv2 = 1 Or Tester.AudioRecordCheck.Value = 1) Then
                  
                      D2K_Release_Card (Card2214)
                     Card2214Exist = 0
                     
                    '=========================================
                    '
                    '   Start Next one song WMA
                    '
                    '========================================
                     Tester.Print "next one press"
                      CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                    
                      Call MsecDelay(0.2)
                      CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                  
                      Tester.Print "next one up"
                        
                        
                         
                      '=========================================
                      '
                      '   End Next one song WMA
                      '
                      '========================================
                    
                     ' Call MsecDelay(1#)
                     '  CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                       Call DACGetData2ChannelUpTriggle
                       Call ADCBuffer2AudioBuffer(1)
                       Call CompareDataTrace(1)
                           
                       rv3 = WMACompareResult
                 Else
                   rv3 = 4
                 End If
                 Call LabelMenu(3, rv3, rv2)
                  
                 '========================================================
                 '
                 '  Recode Raw data to pattern file
                 '
                 '========================================================
                  
                  
                 If Tester.AudioRecordCheck.Value = 1 Then
                                       
                     Open App.Path & "\MP3Pattern\MP3.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, MP3RawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, MP3RawData(1, i)
                         Next i
                      
    
                    Close #4
                
                
                                  
                     Open App.Path & "\WMAPattern\WMA.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, WMARawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, WMARawData(1, i)
                         Next i
                      
    
                    Close #4
                    
                 End If
                 
                 If rv2 * rv3 = 1 Then
                 
                   ' Tester.Cls
                    Tester.Print "MP3Matchmode="; MP3MatchIndex
                    Tester.Print "WMAMatchmode="; WMAMatchIndex
                 End If

                 dma2205_4.Show
                 dma2205_4.ShowData 0, 6000
                
                 Card2214Exist = 0
                  D2K_Release_Card (Card2214)
                  
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
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
                ElseIf rv0 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub
Public Sub AU3150NLF27_1_SlotTest_Noise()
Dim i As Integer
  
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                 '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &H80)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                 CardResult = DO_WritePort(card, Channel_P1A, &H2F)   ' usb Power Off
                 Call MsecDelay(0.8)    'power on time
                  
                 Call MsecDelay(2.5)     'power on time
                  
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
            
                ClosePipe
                LBA = LBA + 1
                OldLBa = LBA
                LBA = LBA + 40000 ' 20000 mp3 shift
                
                rv0 = CBWTest_New_31303(0, 1, "vid_058f")
                LBA = OldLBa
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                Tester.Cls
                Tester.Print "rv0="; rv0
                 
                 
                If rv0 = 1 Then
                    
                          '=================================
                          '   begin load pattern
                          '=================================
                          
                         
                          If OldChipName <> ChipName Then
                                 
                                Call LoadAU3150JLF27AudioPattern
                                OldChipName = ChipName
                          End If
                                                      
              
                          
                          '==========================================
                          '  do a low-high-low pulse at play button
                          '==========================================
                         
                            CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                             
                             Call MsecDelay(0.5)
                             
                             CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                               
                              Call MsecDelay(4)    'power on time and DAC initial time
                               
                             Call MsecDelay(2.5)     'for 3150K power on time

                            CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                          '==========================================
                          '  Get ADC Data and compare result
                          '==========================================
                   Call MsecDelay(1.5)
                           Call DACGetData2ChannelUpTriggle
                           Call ADCBuffer2AudioBuffer(0)
                           Call CompareDataTrace(0)
                           
                           rv2 = MP3CompareResult
                  Else
                   rv2 = 4
                 End If
                 
                 Call LabelMenu(2, rv2, rv0)
                 

               
                If (rv0 = 1) And (rv2 = 1 Or Tester.AudioRecordCheck.Value = 1) Then
                  
                      D2K_Release_Card (Card2214)
                     Card2214Exist = 0
                     
                    '=========================================
                    '
                    '   Start Next one song WMA
                    '
                    '========================================
                     Tester.Print "next one press"
                      CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                    
                      Call MsecDelay(0.2)
                      CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                  
                      Tester.Print "next one up"
                        
                        
                         
                      '=========================================
                      '
                      '   End Next one song WMA
                      '
                      '========================================
                    
                     ' Call MsecDelay(1#)
                     '  CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                       Call DACGetData2ChannelUpTriggle
                       Call ADCBuffer2AudioBuffer(1)
                       Call CompareDataTrace(1)
                           
                       rv3 = WMACompareResult
                 Else
                   rv3 = 4
                 End If
                 Call LabelMenu(3, rv3, rv2)
                  
                 '========================================================
                 '
                 '  Recode Raw data to pattern file
                 '
                 '========================================================
                  
                  
                 If Tester.AudioRecordCheck.Value = 1 Then
                                       
                     Open App.Path & "\MP3Pattern\MP3.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, MP3RawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, MP3RawData(1, i)
                         Next i
                      
    
                    Close #4
                
                
                                  
                     Open App.Path & "\WMAPattern\WMA.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, WMARawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, WMARawData(1, i)
                         Next i
                      
    
                    Close #4
                    
                 End If
                 
                 If rv2 * rv3 = 1 Then
                 
                   ' Tester.Cls
                    Tester.Print "MP3Matchmode="; MP3MatchIndex
                    Tester.Print "WMAMatchmode="; WMAMatchIndex
                 End If

                 dma2205_4.Show
                 dma2205_4.ShowData 0, 6000
                
                 Card2214Exist = 0
                  D2K_Release_Card (Card2214)
                  
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
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
                ElseIf rv0 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub
Public Sub AU3150QLF27_1_SlotTest_Noise()
Dim i As Integer
  
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                 '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &H80)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                 CardResult = DO_WritePort(card, Channel_P1A, &H2F)   ' usb Power Off
                 Call MsecDelay(0.8)    'power on time
                  
                 Call MsecDelay(2.5)     'power on time
                  
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
            
                ClosePipe
                LBA = LBA + 1
                OldLBa = LBA
                LBA = LBA + 40000 ' 20000 mp3 shift
                
                rv0 = CBWTest_New_31303(0, 1, "vid_058f")
                LBA = OldLBa
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                Tester.Cls
                Tester.Print "rv0="; rv0
                 
                 
                If rv0 = 1 Then
                    
                          '=================================
                          '   begin load pattern
                          '=================================
                          
                         
                          If OldChipName <> ChipName Then
                                 
                                Call LoadAU3150JLF27AudioPattern
                                OldChipName = ChipName
                          End If
                                                      
              
                          
                          '==========================================
                          '  do a low-high-low pulse at play button
                          '==========================================
                         
                            CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                             
                             Call MsecDelay(0.5)
                             
                             CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                               
                              Call MsecDelay(4)    'power on time and DAC initial time
                               
                             Call MsecDelay(2.5)     'for 3150K power on time

                            CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                          '==========================================
                          '  Get ADC Data and compare result
                          '==========================================
                   Call MsecDelay(1.5)
                           Call DACGetData2ChannelUpTriggle
                           Call ADCBuffer2AudioBuffer(0)
                           Call CompareDataTrace(0)
                           
                           rv2 = MP3CompareResult
                  Else
                   rv2 = 4
                 End If
                 
                 Call LabelMenu(2, rv2, rv0)
                 

               
                If (rv0 = 1) And (rv2 = 1 Or Tester.AudioRecordCheck.Value = 1) Then
                  
                      D2K_Release_Card (Card2214)
                     Card2214Exist = 0
                     
                    '=========================================
                    '
                    '   Start Next one song WMA
                    '
                    '========================================
                     Tester.Print "next one press"
                      CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                    
                      Call MsecDelay(0.2)
                      CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                  
                      Tester.Print "next one up"
                        
                        
                         
                      '=========================================
                      '
                      '   End Next one song WMA
                      '
                      '========================================
                    
                     ' Call MsecDelay(1#)
                     '  CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                       Call DACGetData2ChannelUpTriggle
                       Call ADCBuffer2AudioBuffer(1)
                       Call CompareDataTrace(1)
                           
                       rv3 = WMACompareResult
                 Else
                   rv3 = 4
                 End If
                 Call LabelMenu(3, rv3, rv2)
                  
                 '========================================================
                 '
                 '  Recode Raw data to pattern file
                 '
                 '========================================================
                  
                  
                 If Tester.AudioRecordCheck.Value = 1 Then
                                       
                     Open App.Path & "\MP3Pattern\MP3.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, MP3RawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, MP3RawData(1, i)
                         Next i
                      
    
                    Close #4
                
                
                                  
                     Open App.Path & "\WMAPattern\WMA.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, WMARawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, WMARawData(1, i)
                         Next i
                      
    
                    Close #4
                    
                 End If
                 
                 If rv2 * rv3 = 1 Then
                 
                   ' Tester.Cls
                    Tester.Print "MP3Matchmode="; MP3MatchIndex
                    Tester.Print "WMAMatchmode="; WMAMatchIndex
                 End If

                 dma2205_4.Show
                 dma2205_4.ShowData 0, 6000
                
                 Card2214Exist = 0
                  D2K_Release_Card (Card2214)
                  
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
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
                ElseIf rv0 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub
Public Sub AU3150PLF27_1_SlotTest_Noise()
Dim i As Integer
  
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                 '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &H80)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                 CardResult = DO_WritePort(card, Channel_P1A, &H2F)   ' usb Power Off
                 Call MsecDelay(0.8)    'power on time
                  
                 Call MsecDelay(2.5)     'power on time
                  
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
            
                ClosePipe
                LBA = LBA + 1
                OldLBa = LBA
                LBA = LBA + 40000 ' 20000 mp3 shift
                
                rv0 = CBWTest_New_31303(0, 1, "vid_058f")
                LBA = OldLBa
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                Tester.Cls
                Tester.Print "rv0="; rv0
                 
                 
                If rv0 = 1 Then
                    
                          '=================================
                          '   begin load pattern
                          '=================================
                          
                         
                          If OldChipName <> ChipName Then
                                 
                                Call LoadAU3150JLF27AudioPattern
                                OldChipName = ChipName
                          End If
                                                      
              
                          
                          '==========================================
                          '  do a low-high-low pulse at play button
                          '==========================================
                         
                            CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                             
                             Call MsecDelay(0.5)
                             
                             CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                               
                              Call MsecDelay(4)    'power on time and DAC initial time
                               
                             Call MsecDelay(2.5)     'for 3150K power on time

                            CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                          '==========================================
                          '  Get ADC Data and compare result
                          '==========================================
                   Call MsecDelay(1.5)
                           Call DACGetData2ChannelUpTriggle
                           Call ADCBuffer2AudioBuffer(0)
                           Call CompareDataTrace(0)
                           
                           rv2 = MP3CompareResult
                  Else
                   rv2 = 4
                 End If
                 
                 Call LabelMenu(2, rv2, rv0)
                 

               
                If (rv0 = 1) And (rv2 = 1 Or Tester.AudioRecordCheck.Value = 1) Then
                  
                      D2K_Release_Card (Card2214)
                     Card2214Exist = 0
                     
                    '=========================================
                    '
                    '   Start Next one song WMA
                    '
                    '========================================
                     Tester.Print "next one press"
                      CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                    
                      Call MsecDelay(0.2)
                      CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                  
                      Tester.Print "next one up"
                        
                        
                         
                      '=========================================
                      '
                      '   End Next one song WMA
                      '
                      '========================================
                    
                     ' Call MsecDelay(1#)
                     '  CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                       Call DACGetData2ChannelUpTriggle
                       Call ADCBuffer2AudioBuffer(1)
                       Call CompareDataTrace(1)
                           
                       rv3 = WMACompareResult
                 Else
                   rv3 = 4
                 End If
                 Call LabelMenu(3, rv3, rv2)
                  
                 '========================================================
                 '
                 '  Recode Raw data to pattern file
                 '
                 '========================================================
                  
                  
                 If Tester.AudioRecordCheck.Value = 1 Then
                                       
                     Open App.Path & "\MP3Pattern\MP3.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, MP3RawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, MP3RawData(1, i)
                         Next i
                      
    
                    Close #4
                
                
                                  
                     Open App.Path & "\WMAPattern\WMA.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, WMARawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, WMARawData(1, i)
                         Next i
                      
    
                    Close #4
                    
                 End If
                 
                 If rv2 * rv3 = 1 Then
                 
                   ' Tester.Cls
                    Tester.Print "MP3Matchmode="; MP3MatchIndex
                    Tester.Print "WMAMatchmode="; WMAMatchIndex
                 End If

                 dma2205_4.Show
                 dma2205_4.ShowData 0, 6000
                
                 Card2214Exist = 0
                  D2K_Release_Card (Card2214)
                  
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
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
                ElseIf rv0 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub
Public Sub AU3150JLF27_1_SlotTest_Noise()
Dim i As Integer
  
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                 '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &H80)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                 CardResult = DO_WritePort(card, Channel_P1A, &H2F)   ' usb Power Off
                 Call MsecDelay(0.8)    'power on time
                  
                 Call MsecDelay(2.5)     'power on time
                  
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
            
                ClosePipe
                LBA = LBA + 1
                OldLBa = LBA
                
                LBA = LBA + 40000 ' 20000 mp3 shift
                
                rv0 = CBWTest_New_31303(0, 1, "vid_058f")
                LBA = OldLBa
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                Tester.Cls
                Tester.Print "rv0="; rv0
                 
                 
                If rv0 = 1 Then
                    
                          '=================================
                          '   begin load pattern
                          '=================================
                          
                         
                          If OldChipName <> ChipName Then
                                 
                                Call LoadAU3150JLF27AudioPattern
                                OldChipName = ChipName
                          End If
                                                      
              
                          
                          '==========================================
                          '  do a low-high-low pulse at play button
                          '==========================================
                         
                            CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                             
                             Call MsecDelay(0.5)
                             
                             CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                               
                              Call MsecDelay(4)    'power on time and DAC initial time
                               
                             Call MsecDelay(2.5)     'for 3150K power on time

                            CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                          '==========================================
                          '  Get ADC Data and compare result
                          '==========================================
                   Call MsecDelay(1.5)
                           Call DACGetData2ChannelUpTriggle
                           Call ADCBuffer2AudioBuffer(0)
                           Call CompareDataTrace(0)
                           
                           rv2 = MP3CompareResult
                  Else
                   rv2 = 4
                 End If
                 
                 Call LabelMenu(2, rv2, rv0)
                 

               
                If (rv0 = 1) And (rv2 = 1 Or Tester.AudioRecordCheck.Value = 1) Then
                  
                      D2K_Release_Card (Card2214)
                     Card2214Exist = 0
                     
                    '=========================================
                    '
                    '   Start Next one song WMA
                    '
                    '========================================
                     Tester.Print "next one press"
                      CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                    
                      Call MsecDelay(0.2)
                      CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                  
                      Tester.Print "next one up"
                        
                        
                         
                      '=========================================
                      '
                      '   End Next one song WMA
                      '
                      '========================================
                    
                     ' Call MsecDelay(1#)
                     '  CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                       Call DACGetData2ChannelUpTriggle
                       Call ADCBuffer2AudioBuffer(1)
                       Call CompareDataTrace(1)
                           
                       rv3 = WMACompareResult
                 Else
                   rv3 = 4
                 End If
                 Call LabelMenu(3, rv3, rv2)
                  
                 '========================================================
                 '
                 '  Recode Raw data to pattern file
                 '
                 '========================================================
                  
                  
                 If Tester.AudioRecordCheck.Value = 1 Then
                                       
                     Open App.Path & "\MP3Pattern\MP3.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, MP3RawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, MP3RawData(1, i)
                         Next i
                      
    
                    Close #4
                
                
                                  
                     Open App.Path & "\WMAPattern\WMA.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, WMARawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, WMARawData(1, i)
                         Next i
                      
    
                    Close #4
                    
                 End If
                 
                 If rv2 * rv3 = 1 Then
                 
                   ' Tester.Cls
                    Tester.Print "MP3Matchmode="; MP3MatchIndex
                    Tester.Print "WMAMatchmode="; WMAMatchIndex
                 End If

                 dma2205_4.Show
                 dma2205_4.ShowData 0, 6000
                
                 Card2214Exist = 0
                  D2K_Release_Card (Card2214)
                  
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
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
                ElseIf rv0 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub

 

Public Sub AU3150ALF28_2_SlotTest_Noise()
Dim i As Integer
  
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                 '=========================================
                '    POWER on
                '=========================================
                 '   CardResult = DO_WritePort(card, Channel_P1A, &H80)
                    CardResult = DO_WritePort(card, Channel_P1A, &HC0)  ' Allen debug
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                CardResult = DO_WritePort(card, Channel_P1A, &H6F)
              '    CardResult = DO_WritePort(card, Channel_P1A, &H2F)   ' usb Power Off
                  Call MsecDelay(0.8)    'power on time
                  
              '   Call MsecDelay(1#)      'power on time
                  
                  
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
            
                ClosePipe
                LBA = LBA + 1
                OldLBa = LBA
                LBA = LBA + 40000 ' 20000 mp3 shift
                
                rv0 = CBWTest_New_31303(0, 1, "vid_058f")
                LBA = OldLBa
                   
               
                 Call LabelMenu(0, rv0, 1)
                 ClosePipe
                 rv1 = CBWTest_New_31302(1, rv0, "vid_058f")
                 Call LabelMenu(1, rv1, rv0)
                 ClosePipe
                 
                 Tester.Print " Flash R/W test rv0="; rv0
                 
                 Tester.Print "SD R/W teswt rv1="; rv1
                 
                If rv0 * rv1 = 1 Then
                
                          '=================================
                          '   begin load pattern
                          '=================================
                          
                         
                          If OldChipName <> ChipName Then
                                 
                                Call LoadAU3150JLF27AudioPattern
                                OldChipName = ChipName
                          End If
                                                      
              
                          
                            '==========================================
                          '  do a low-high-low pulse at play button
                          '==========================================
                         
                             
                             
                             CardResult = DO_WritePort(card, Channel_P1A, &HCF)
                             
                             Call MsecDelay(0.3)
                             CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                                 Call MsecDelay(2.5)
                                'Call MsecDelay(4.5)
                              CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                            '   CardResult = DO_WritePort(card, Channel_P1A, &HDF)
                          '    Call MsecDelay(4)    'power on time and DAC initial time
                               
                            ' Call MsecDelay(2.5)     'for 3150K power on time
                               Call MsecDelay(0.3)
                             CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                            
                            ' CardResult = DO_WritePort(card, Channel_P1A, &HCF)
                             
                          '==========================================
                          '  Get ADC Data and compare result
                          '==========================================
                   Call MsecDelay(1.5)
                           Call DACGetData2ChannelUpTriggle
                           Call ADCBuffer2AudioBuffer(0)
                           Call CompareDataTrace(0)
                           
                           rv2 = MP3CompareResult
                  Else
                   rv2 = 4
                 End If
                 
                 Call LabelMenu(2, rv2, rv0)
                 

               
                If (rv0 * rv1 = 1) And (rv2 = 1 Or Tester.AudioRecordCheck.Value = 1) Then
                  
                      D2K_Release_Card (Card2214)
                     Card2214Exist = 0
                     
                    '=========================================
                    '
                    '   Start Next one song WMA
                    '
                    '========================================
                     Tester.Print "next one press"
                      CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                    
                      Call MsecDelay(0.2)
                      CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                  
                      Tester.Print "next one up"
                        
                        
                         
                      '=========================================
                      '
                      '   End Next one song WMA
                      '
                      '========================================
                    
                     ' Call MsecDelay(1#)
                     '  CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                       Call DACGetData2ChannelUpTriggle
                       Call ADCBuffer2AudioBuffer(1)
                       Call CompareDataTrace(1)
                           
                       rv3 = WMACompareResult
                 Else
                   rv3 = 4
                 End If
                 Call LabelMenu(3, rv3, rv2)
                  
                 '========================================================
                 '
                 '  Recode Raw data to pattern file
                 '
                 '========================================================
                  
                  
                 If Tester.AudioRecordCheck.Value = 1 Then
                                       
                     Open App.Path & "\MP3Pattern\MP3.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, MP3RawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, MP3RawData(1, i)
                         Next i
                      
    
                    Close #4
                
                
                                  
                     Open App.Path & "\WMAPattern\WMA.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, WMARawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, WMARawData(1, i)
                         Next i
                      
    
                    Close #4
                    
                 End If
                 
                 If rv2 * rv3 = 1 Then
                 
                   ' Tester.Cls
                    Tester.Print "MP3Matchmode="; MP3MatchIndex
                    Tester.Print "WMAMatchmode="; WMAMatchIndex
                 End If

                 dma2205_4.Show
                 dma2205_4.ShowData 0, 6000
                
                 Card2214Exist = 0
                  D2K_Release_Card (Card2214)
                  
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
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
                ElseIf rv0 * rv1 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub
Public Sub AU3150ALF27_2_SlotTest_Noise()
Dim i As Integer
  
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                 '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &H80)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                 CardResult = DO_WritePort(card, Channel_P1A, &H2F)   ' usb Power Off
                 Call MsecDelay(0.8)    'power on time
                  
                 Call MsecDelay(2.5)     'power on time
                  
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
            
                ClosePipe
                LBA = LBA + 1
                OldLBa = LBA
                LBA = LBA + 40000 ' 20000 mp3 shift
                
                rv0 = CBWTest_New_31303(0, 1, "vid_058f")
                LBA = OldLBa
                   
               
                 Call LabelMenu(0, rv0, 1)
                 ClosePipe
                 rv1 = CBWTest_New_31302(1, rv0, "vid_058f")
                 Call LabelMenu(1, rv1, rv0)
                 ClosePipe
                 
                 Tester.Print " Flash R/W test rv0="; rv0
                 
                 Tester.Print "SD R/W teswt rv1="; rv1
                 
                If rv0 * rv1 = 1 Then
                
                          '=================================
                          '   begin load pattern
                          '=================================
                          
                         
                          If OldChipName <> ChipName Then
                                 
                                Call LoadAU3150JLF27AudioPattern
                                OldChipName = ChipName
                          End If
                                                      
              
                          
                          '==========================================
                          '  do a low-high-low pulse at play button
                          '==========================================
                         
                            CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                             
                             Call MsecDelay(0.5)
                             
                             CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                               
                              Call MsecDelay(4)    'power on time and DAC initial time
                               
                             Call MsecDelay(2.5)     'for 3150K power on time

                            CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                          '==========================================
                          '  Get ADC Data and compare result
                          '==========================================
                   Call MsecDelay(1.5)
                           Call DACGetData2ChannelUpTriggle
                           Call ADCBuffer2AudioBuffer(0)
                           Call CompareDataTrace(0)
                           
                           rv2 = MP3CompareResult
                  Else
                   rv2 = 4
                 End If
                 
                 Call LabelMenu(2, rv2, rv0)
                 

               
                If (rv0 * rv1 = 1) And (rv2 = 1 Or Tester.AudioRecordCheck.Value = 1) Then
                  
                      D2K_Release_Card (Card2214)
                     Card2214Exist = 0
                     
                    '=========================================
                    '
                    '   Start Next one song WMA
                    '
                    '========================================
                     Tester.Print "next one press"
                      CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                    
                      Call MsecDelay(0.2)
                      CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                  
                      Tester.Print "next one up"
                        
                        
                         
                      '=========================================
                      '
                      '   End Next one song WMA
                      '
                      '========================================
                    
                     ' Call MsecDelay(1#)
                     '  CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                       Call DACGetData2ChannelUpTriggle
                       Call ADCBuffer2AudioBuffer(1)
                       Call CompareDataTrace(1)
                           
                       rv3 = WMACompareResult
                 Else
                   rv3 = 4
                 End If
                 Call LabelMenu(3, rv3, rv2)
                  
                 '========================================================
                 '
                 '  Recode Raw data to pattern file
                 '
                 '========================================================
                  
                  
                 If Tester.AudioRecordCheck.Value = 1 Then
                                       
                     Open App.Path & "\MP3Pattern\MP3.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, MP3RawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, MP3RawData(1, i)
                         Next i
                      
    
                    Close #4
                
                
                                  
                     Open App.Path & "\WMAPattern\WMA.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, WMARawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, WMARawData(1, i)
                         Next i
                      
    
                    Close #4
                    
                 End If
                 
                 If rv2 * rv3 = 1 Then
                 
                   ' Tester.Cls
                    Tester.Print "MP3Matchmode="; MP3MatchIndex
                    Tester.Print "WMAMatchmode="; WMAMatchIndex
                 End If

                 dma2205_4.Show
                 dma2205_4.ShowData 0, 6000
                
                 Card2214Exist = 0
                  D2K_Release_Card (Card2214)
                  
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
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
                ElseIf rv0 * rv1 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub
 
Public Sub AU3150CLF27_2_SlotTest_Noise()
Dim i As Integer
  
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                 '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &H80)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                 CardResult = DO_WritePort(card, Channel_P1A, &H2F)   ' usb Power Off
                 Call MsecDelay(0.8)    'power on time
                  
                 Call MsecDelay(2.5)     'power on time
                  
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
            
                ClosePipe
                LBA = LBA + 1
                OldLBa = LBA
                LBA = LBA + 40000 ' 20000 mp3 shift
                
                rv0 = CBWTest_New_31303(0, 1, "vid_058f")
                LBA = OldLBa
                   
               
                 Call LabelMenu(0, rv0, 1)
                 ClosePipe
                 rv1 = CBWTest_New_31302(1, rv0, "vid_058f")
                 Call LabelMenu(1, rv1, rv0)
                 ClosePipe
                 
                 Tester.Print " Flash R/W test rv0="; rv0
                 
                 Tester.Print "SD R/W teswt rv1="; rv1
                 
                If rv0 * rv1 = 1 Then
                
                          '=================================
                          '   begin load pattern
                          '=================================
                          
                         
                          If OldChipName <> ChipName Then
                                 
                                Call LoadAU3150JLF27AudioPattern
                                OldChipName = ChipName
                          End If
                                                      
              
                          
                          '==========================================
                          '  do a low-high-low pulse at play button
                          '==========================================
                         
                            CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                             
                             Call MsecDelay(0.5)
                             
                             CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                               
                              Call MsecDelay(4)    'power on time and DAC initial time
                               
                             Call MsecDelay(2.5)     'for 3150K power on time

                            CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                          '==========================================
                          '  Get ADC Data and compare result
                          '==========================================
                   Call MsecDelay(1.5)
                           Call DACGetData2ChannelUpTriggle
                           Call ADCBuffer2AudioBuffer(0)
                           Call CompareDataTrace(0)
                           
                           rv2 = MP3CompareResult
                  Else
                   rv2 = 4
                 End If
                 
                 Call LabelMenu(2, rv2, rv0)
                 

               
                If (rv0 * rv1 = 1) And (rv2 = 1 Or Tester.AudioRecordCheck.Value = 1) Then
                  
                      D2K_Release_Card (Card2214)
                     Card2214Exist = 0
                     
                    '=========================================
                    '
                    '   Start Next one song WMA
                    '
                    '========================================
                     Tester.Print "next one press"
                      CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                    
                      Call MsecDelay(0.2)
                      CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                  
                      Tester.Print "next one up"
                        
                        
                         
                      '=========================================
                      '
                      '   End Next one song WMA
                      '
                      '========================================
                    
                     ' Call MsecDelay(1#)
                     '  CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                       Call DACGetData2ChannelUpTriggle
                       Call ADCBuffer2AudioBuffer(1)
                       Call CompareDataTrace(1)
                           
                       rv3 = WMACompareResult
                 Else
                   rv3 = 4
                 End If
                 Call LabelMenu(3, rv3, rv2)
                  
                 '========================================================
                 '
                 '  Recode Raw data to pattern file
                 '
                 '========================================================
                  
                  
                 If Tester.AudioRecordCheck.Value = 1 Then
                                       
                     Open App.Path & "\MP3Pattern\MP3.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, MP3RawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, MP3RawData(1, i)
                         Next i
                      
    
                    Close #4
                
                
                                  
                     Open App.Path & "\WMAPattern\WMA.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, WMARawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, WMARawData(1, i)
                         Next i
                      
    
                    Close #4
                    
                 End If
                 
                 If rv2 * rv3 = 1 Then
                 
                   ' Tester.Cls
                    Tester.Print "MP3Matchmode="; MP3MatchIndex
                    Tester.Print "WMAMatchmode="; WMAMatchIndex
                 End If

                 dma2205_4.Show
                 dma2205_4.ShowData 0, 6000
                
                 Card2214Exist = 0
                  D2K_Release_Card (Card2214)
                  
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
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
                ElseIf rv0 * rv1 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub
Public Sub AU3150CLF2C_2_SlotTest_Noise()
Dim i As Integer
  
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                 '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &H80)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                 CardResult = DO_WritePort(card, Channel_P1A, &H2F)   ' usb Power Off
                 Call MsecDelay(0.8)    'power on time
                  
                 Call MsecDelay(2.5)     'power on time
                  
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
            
                ClosePipe
                LBA = LBA + 1
                OldLBa = LBA
                LBA = LBA + 40000 ' 20000 mp3 shift
                
                rv0 = CBWTest_New_31303(0, 1, "vid_058f")
                LBA = OldLBa
                   
               
                 Call LabelMenu(0, rv0, 1)
                 ClosePipe
                 rv1 = CBWTest_New_31302(1, rv0, "vid_058f")
                 Call LabelMenu(1, rv1, rv0)
                 ClosePipe
                 
                 Tester.Print " Flash R/W test rv0="; rv0
                 
                 Tester.Print "SD R/W teswt rv1="; rv1
                 
                If rv0 * rv1 = 1 Then
                
                           '=================================
                          '   begin load pattern
                          '=================================
                          
                         
                          If OldChipName <> ChipName Then
                                 
                              '  Call LoadAU3150JLF27AudioPattern
                                OldChipName = ChipName
                          End If
                                                      
              
                          
                          '==========================================
                          '  do a low-high-low pulse at play button
                          '==========================================
                         
                            CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                             
                             Call MsecDelay(0.5)
                             
                             CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                               
                              Call MsecDelay(4)    'power on time and DAC initial time
                               
                             Call MsecDelay(2.5)     'for 3150K power on time

                            CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                          '==========================================
                          '  Get ADC Data and compare result
                          '==========================================
                   Call MsecDelay(1.5)
                           Call DACGetData2ChannelUpTriggle
                           Call ADCBuffer2AudioBuffer(0)
                            Call DistingDataTrace(0)
                           rv2 = MP3DistinguishResult
                  Else
                   rv2 = 4
                 End If
                 
                 Call LabelMenu(2, rv2, rv0)
                 

               
                If (rv0 * rv1 = 1) And (rv2 = 1 Or Tester.AudioRecordCheck.Value = 1) Then
                  
                      D2K_Release_Card (Card2214)
                     Card2214Exist = 0
                     
                    '=========================================
                    '
                    '   Start Next one song WMA
                    '
                    '========================================
                     Tester.Print "next one press"
                      CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                    
                      Call MsecDelay(0.2)
                      CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                  
                      Tester.Print "next one up"
                        
                        
                         
                      '=========================================
                      '
                      '   End Next one song WMA
                      '
                      '========================================
                    
                     ' Call MsecDelay(1#)
                     '  CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                       Call DACGetData2ChannelUpTriggle
                       Call ADCBuffer2AudioBuffer(1)
                         Call DistingDataTrace(1)
                           rv3 = WMADistinguishResult
                 Else
                   rv3 = 4
                 End If
                 Call LabelMenu(3, rv3, rv2)
                  
                 '========================================================
                 '
                 '  Recode Raw data to pattern file
                 '
                 '========================================================
                  
                  
                 If Tester.AudioRecordCheck.Value = 1 Then
                                       
                     Open App.Path & "\MP3Pattern\MP3.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, MP3RawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, MP3RawData(1, i)
                         Next i
                      
    
                    Close #4
                
                
                                  
                     Open App.Path & "\WMAPattern\WMA.txt" For Output As #4
    
                         For i = 0 To 5999
                           Print #4, WMARawData(0, i)
                         Next i
                          For i = 0 To 5999
                           Print #4, WMARawData(1, i)
                         Next i
                      
    
                    Close #4
                    
                 End If
                 
                 If rv2 * rv3 = 1 Then
                 
                   ' Tester.Cls
                    Tester.Print "MP3Matchmode="; MP3MatchIndex
                    Tester.Print "WMAMatchmode="; WMAMatchIndex
                 End If

                 dma2205_4.Show
                 dma2205_4.ShowData3 0, 6000
                
                 Card2214Exist = 0
                  D2K_Release_Card (Card2214)
                  
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
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
                ElseIf rv0 * rv1 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub

Public Sub AU3152ALF24_1_SlotTest_Noise()
Dim i As Integer
  
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                 '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &H80)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                 CardResult = DO_WritePort(card, Channel_P1A, &H2F)   ' usb Power Off
                 Call MsecDelay(0.8)    'power on time
                  
                 Call MsecDelay(2.5)     'power on time
                  
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
            
                ClosePipe
                LBA = LBA + 1
                OldLBa = LBA
                LBA = LBA + 20000 ' 20000 mp3 shift
                
                rv0 = CBWTest_New_31303(0, 1, "vid_058f")
                LBA = OldLBa
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                Tester.Cls
                Tester.Print "rv0="; rv0
                
                
                 
                If rv0 = 1 Then
                    
                          '=================================
                          '   begin load pattern
                          '=================================
                          
                         
                          If OldChipName <> ChipName Then
                                 
                                Call LoadAU3152ALF24AudioPattern
                                OldChipName = ChipName
                          End If
                                                      
              
                          
                          '==========================================
                          '  do a low-high-low pulse at play button
                          '==========================================
                         
                            CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                             
                             Call MsecDelay(0.5)
                             
                             CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                               
                             Call MsecDelay(4)    'power on time and DAC initial time
                               
                             Call MsecDelay(2.5)     'for 3150K power on time

                            CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                          '==========================================
                          '  Get ADC Data and compare result
                          '==========================================
                  
                           Call DACGetData2ChannelLowTriggle64
                           Call ADCBuffer2AudioBuffer(0)
                           Call CompareData(0)
                           
                           rv2 = MP3CompareResult
                  Else
                   rv2 = 4
                 End If
                 
                 Call LabelMenu(2, rv2, rv0)
                 

               
                If rv2 = 1 Or Tester.AudioRecordCheck.Value = 1 Then
                  
                      D2K_Release_Card (Card2214)
                     Card2214Exist = 0
                     
                    '=========================================
                    '
                    '   Start Next one song WMA
                    '
                    '========================================
                     Tester.Print "next one press"
                      CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                    
                      Call MsecDelay(0.5)
                      CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                  
                      Tester.Print "next one up"
                        
                        
                         
                      '=========================================
                      '
                      '   End Next one song WMA
                      '
                      '========================================
                    
                 
                       Call DACGetData2ChannelLowTriggle64
                       Call ADCBuffer2AudioBuffer(1)
                       Call CompareData(1)
                           
                       rv3 = WMACompareResult
                 Else
                   rv3 = 4
                 End If
                 Call LabelMenu(3, rv3, rv2)
                  
                 '========================================================
                 '
                 '  Recode Raw data to pattern file
                 '
                 '========================================================
                  
                  
                 If Tester.AudioRecordCheck.Value = 1 Then
                                       
                     Open App.Path & "\MP3Pattern\MP3.txt" For Output As #4
    
                         For i = 0 To 99
                           Print #4, MP3RawData(0, i)
                         Next i
                          For i = 0 To 99
                           Print #4, MP3RawData(1, i)
                         Next i
                      
    
                    Close #4
                
                
                                  
                     Open App.Path & "\WMAPattern\WMA.txt" For Output As #4
    
                         For i = 0 To 99
                           Print #4, WMARawData(0, i)
                         Next i
                          For i = 0 To 99
                           Print #4, WMARawData(1, i)
                         Next i
                      
    
                    Close #4
                    
                 End If
                 
                 If rv2 * rv3 = 1 Then
                 
                   ' Tester.Cls
                    Tester.Print "MP3Matchmode="; MP3MatchIndex
                    Tester.Print "WMAMatchmode="; WMAMatchIndex
                 End If

                 dma2205_4.Show
                 dma2205_4.ShowData 0, 100
                
                 Card2214Exist = 0
                  D2K_Release_Card (Card2214)
                  
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
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
                ElseIf rv0 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub


Public Sub AU3130_1_SlotTest_WMA()
Dim i As Integer

   If ChipName = "AU3150ALF21" Or ChipName = "AU3150ILF21" Or ChipName = "AU3150CLF21" Or ChipName = "AU3150LLF21" Then
                   ChipName = "AU3150JLF21"
              End If
           
   If Left(ChipName, 10) = "AU3150KLF2" Then
                   TmpChip = ChipName
                   ChipName = "AU3150JLF21"
    End If
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                
                
                 '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &H80)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                 CardResult = DO_WritePort(card, Channel_P1A, &H2F)   ' usb Power Off
                 Call MsecDelay(0.8)    'power on time
                 
                 If Left(TmpChip, 10) = "AU3150KLF2" Then
                   Call MsecDelay(2.5)     'power on time
                 End If
                 
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
                 
                
                
                ClosePipe
                LBA = LBA + 1
                OldLBa = LBA
                LBA = LBA + 20000 ' 20000 mp3 shift
                
               
                If Left(TmpChip, 10) = "AU3150KLF2" Then
                   rv0 = CBWTest_New_31303(0, 1, "vid_058f")
                 Else
                   rv0 = CBWTest_New_3130(0, 1, "vid_058f")
                 
                 End If
                
                LBA = OldLBa
                'Print "a1"
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 
                 
                If rv0 = 1 Then
                
                        ' =========== save MP3 pattern
                                If ChipName = "AU3130" Then
                                For i = 0 To 100
                                   MP3Data(i) = MP3Data_A(i)
                                Next
                            End If
                            
                             If ChipName = "AU3130B" Or ChipName = "AU3130CL" Then
                                For i = 0 To 100
                                   MP3Data(i) = MP3Data_B(i)
                                   MP3Data_1(i) = MP3Data_B1(i)
                                   MP3Data_2(i) = MP3Data_B2(i)
                                   MP3Data_3(i) = MP3Data_C(i)
                                   MP3Data_4(i) = MP3Data_C1(i)
                                   MP3Data_5(i) = MP3Data_C2(i)
                                  
                                Next
                            End If
                            
                               If ChipName = "AU3130BLF20" Then
                                For i = 0 To 100
                                   MP3Data_6(i) = MP3Data_BL(i)
                                   MP3Data_7(i) = MP3Data_BL1(i)
                                Next
                            End If
                            
                             If ChipName = "AU3130CLF20" Then
                                For i = 0 To 100
                                   MP3Data_8(i) = MP3Data_CL(i)
                                   
                                Next
                            End If
                            
                            '(4)============================2007.06.14 CHEYENNE CHANG
                           If ChipName = "AU3130CLF20" Then
                                For i = 0 To 100
                                   MP3Data_9(i) = MP3Data_CW1(i)
                                   
                                Next
                            End If

                            
                            If ChipName = "AU3150JLF21" Then
                                For i = 0 To 100
                                   MP3Data_10(i) = MP3Data_3150J(i)
                                   MP3Data_11(i) = MP3Data_3150J1(i)
                                Next
                            End If
  
                            If ChipName = "AU3152ALF21" Or ChipName = "AU3152ALF23" Then
                                For i = 0 To 100
                                   MP3Data_15(i) = MP3Data_3152A1(i)
                                   MP3Data_16(i) = MP3Data_3152A2(i)
                                   MP3Data_161(i) = MP3Data_3152A3(i)
                                Next
                            End If
                            
                            
                             If TmpChip = "AU3150KLF21" Then
                                For i = 0 To 100
                                   MP3Data_19(i) = MP3Data_3150KL1(i)
                                   MP3Data_20(i) = MP3Data_3150KL2(i)
                                Next
                            End If
                            
                               If TmpChip = "AU3150KLF22" Then
                                For i = 0 To 100
                                   MP3Data_21(i) = MP3Data_3150KL21(i)
                                   MP3Data_22(i) = MP3Data_3150KL22(i)
                                   
                                Next
                            End If
                            
                            
                              If TmpChip = "AU3150KLF23" Then
                                For i = 0 To 100
                                   MP3Data_21(i) = MP3Data_3150KL21(i)
                                   MP3Data_22(i) = MP3Data_3150KL22(i)
                                   MP3Data_23(i) = MP3Data_3150KL23(i)
                                Next
                            End If
                            
                            
                              If TmpChip = "AU3150KLF24" Then
                                For i = 0 To 100
                                   MP3Data_21(i) = MP3Data_3150KL21(i)
                                   MP3Data_22(i) = MP3Data_3150KL22(i)
                                   MP3Data_23(i) = MP3Data_3150KL23(i)
                                   MP3WMA_23(i) = MP3WMA_3150KL23(i)
                                   MP3WMA_231(i) = MP3WMA_3150KL231(i)
                                   MP3WMA_232(i) = MP3WMA_3150KL232(i)
                                Next
                            End If
                          '================================
                              CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                             If CardResult <> 0 Then
                                MsgBox "Power off fail"
                                End
                             End If
                            ' mp3==================================================
                             Call MsecDelay(0.5)
                             
                             CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                               
                             Call MsecDelay(4)    'power on time and DAC initial time
                               If Left(TmpChip, 10) = "AU3150KLF2" Then  'debug
                                 Call MsecDelay(2.5)     'power on time
                              End If
                 
                             
                            ' Data initial =================
           
                            CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
         
                            If Left(ChipName, 7) = "AU3130B" Then    ' MP3 Power Off
                              Call MsecDelay(0.5)
                             
                              CardResult = DO_WritePort(card, Channel_P1A, &H8D)
                           
                            End If
                            
                            
                               If Left(ChipName, 7) = "AU3150J" Or Left(ChipName, 7) = "AU3152A" Then      ' MP3 Power Off
                           '   Call MsecDelay(0.5)
                             
                              CardResult = DO_WritePort(card, Channel_P1A, &H8D)
                              
                           '   Call MsecDelay(10)
                            End If
                            
                            If Left(ChipName, 7) = "AU3130C" Then    ' MP3 Power Off
                              Call MsecDelay(2)
                              CardResult = DO_WritePort(card, Channel_P1A, &H8D)
                            End If
                           
                              If CardResult <> 0 Then
                                MsgBox "Power on fail"
                                End
                             End If
                             
                    rv2 = AU3130TestMP3
                  Else
                   rv2 = 4
                 End If
                 
                 Call LabelMenu(2, rv2, rv0)
                 
                  
                  
                  ' Call MsecDelay(2)
                 
                  

                 '=========================================
                 '
                 '   Start Next one song WMA
                 '
                 '========================================
                
                  CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                  Tester.Print "next one press"
                     Call MsecDelay(0.5)
                   CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                     Call MsecDelay(0.5)
                     CardResult = DO_WritePort(card, Channel_P1A, &H8E)
                
                    Call MsecDelay(0.5)
                   CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                   Tester.Print "next one up"
                   
                   
                     
                   
                  '     CardResult = DO_WritePort(card, Channel_P1A, &H9F)
                 '   Call MsecDelay(1)
                 '  CardResult = DO_WritePort(card, Channel_P1A, &H9E)
                 '  Call MsecDelay(1)
                 
                 '   CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                   ' Call MsecDelay(1)
                   
                   '=========================================
                 '
                 '   End Next one song WMA
                 '
                 '========================================
                  
                 If rv2 = 1 Then
                   rv3 = AU3130TestWMA
                 Else
                   rv3 = 4
                 End If
                 Call LabelMenu(3, rv3, rv2)
                  
                 
                 
                     Open App.Path & "\MP3_3150KL23WMA.txt" For Output As #4

                        For i = 0 To 100
                       Print #4, gnBufferWMA(i)
                      Next i
                  

                      Close #4

                 dma2205_2.Show
                 dma2205_2.ShowData 0, 100
                
                  
                '  dma2205.Hide
                 
                  
                  
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
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
                ElseIf rv0 * rv2 * rv3 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub

Public Sub AU3150MLTestSub()
     
                    If PCI7248InitFinish = 0 Then
                      PCI7248Exist
                    End If
                
                   result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    
                    CardResult = DO_WritePort(card, Channel_P1A, &H1)  ' 1111 1110
                    Call MsecDelay(0.03)
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                    CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' 1111 1110
                    Call MsecDelay(0.7)
                   
                
                  
                
                
                LBA = LBA + 1
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                rv1 = CBWTest_New_31302(1, rv0, "vid_058f")
                 Call LabelMenu(1, rv1, rv0)
                 ClosePipe
                 
               
                If rv0 * rv1 = 1 Then
                     rv2 = 1
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                     If LightOn <> 223 Then
                        rv2 = 2
                     End If
                Else
                    rv2 = 4
                     
                 End If
                         Call LabelMenu(2, rv2, rv1)
                  Tester.Print " Flash R/W test rv0="; rv0
                 
                  Tester.Print "SD R/W teswt rv1="; rv1
                  Tester.Print "GPO    teswt rv2="; rv2
                Tester.Print "LBA="; LBA
            
                    If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
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
                ElseIf rv0 * rv1 * rv2 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub
Function CBWTest_New_31302(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
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
        CBWTest_New_31302 = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New_31302 = 0
   
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
      CBWTest_New_31302 = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_New_31302 = 2   ' Write fail
      Exit Function
    End If
 
    '======================================
    
     ' for unitSpeed
    
    TmpInteger = TestUnitSpeed(Lun)
    
    If TmpInteger = 0 Then
        
       CBWTest_New_31302 = 2   ' usb 2.0 high speed fail
       UsbSpeedTestResult = 2
       Exit Function
    End If
    
   
    
    TmpInteger = TestUnitReady(Lun)
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
           CBWTest_New_31302 = 2  'Write fail
           Exit Function
        End If
        
    End If
    '======================================
     'TmpInteger = Read_CapacityAU6371(LBA, Lun, 8)
    TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
 
        
      
    If TmpInteger = 0 Then
         CBWTest_New_31302 = 2  'write fail
          Exit Function
     End If
    
      
    TmpInteger = Write_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_31302 = 2  'write fail
        Exit Function
    End If
    
    TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_31302 = 3    'Read fail
        Exit Function
    End If
     
    For i = 0 To CBWDataTransferLength - 1
    
        If ReadData(i) <> Pattern(i) Then
          CBWTest_New_31302 = 3    'Read fail
          Exit Function
        End If
    
    Next
    
    CBWTest_New_31302 = 1
        
    
    End Function

Function CBWTest_New_31303(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
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
        CBWTest_New_31303 = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New_31303 = 0
   
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
      CBWTest_New_31303 = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_New_31303 = 2   ' Write fail
      Exit Function
    End If
 
    '======================================
    
     ' for unitSpeed
    
    TmpInteger = TestUnitSpeed(Lun)
    
    If TmpInteger = 0 Then
        
       CBWTest_New_31303 = 2   ' usb 2.0 high speed fail
       UsbSpeedTestResult = 2
       Tester.Print "speed Fail"
       Exit Function
    End If
    
    
    
    TmpInteger = TestUnitReady(Lun)
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
           Tester.Print "request Fail"
           CBWTest_New_31303 = 2  'Write fail
           Exit Function
        End If
        
    End If
    '======================================
   TmpInteger = Read_Data1(LBA, Lun, CBWDataTransferLength)
  
   
      
    TmpInteger = Write_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        Tester.Print "write fail Fail"
        CBWTest_New_31303 = 2  'write fail
        Exit Function
    End If
    
    TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_31303 = 3    'Read fail
        Exit Function
    End If
     
    For i = 0 To CBWDataTransferLength - 1
    
        If ReadData(i) <> Pattern(i) Then
          CBWTest_New_31303 = 3    'Read fail
          Exit Function
        End If
    
    Next
    
    CBWTest_New_31303 = 1
        
    
    End Function

Function CBWTest_New_3130(Lun As Byte, PreSlotStatus As Byte, Vid_PID As String) As Byte
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
        CBWTest_New_3130 = 4
        Exit Function
    End If
    '========================================
   
    CBWTest_New_3130 = 0
   
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
      CBWTest_New_3130 = 0   ' no readerExist
      ReaderExist = 0
      Exit Function
    End If
    '=======================================
    If OpenPipe = 0 Then
      CBWTest_New_3130 = 2   ' Write fail
      Exit Function
    End If
 
    '======================================
    
     ' for unitSpeed
    
    TmpInteger = TestUnitSpeed(Lun)
    
    If TmpInteger = 0 Then
        
       CBWTest_New_3130 = 2   ' usb 2.0 high speed fail
       UsbSpeedTestResult = 2
       Exit Function
    End If
    
    
    
    TmpInteger = TestUnitReady(Lun)
    If TmpInteger = 0 Then
        TmpInteger = RequestSense(Lun)
        
        If TmpInteger = 0 Then
        
           CBWTest_New_3130 = 2  'Write fail
           Exit Function
        End If
        
    End If
    '======================================
   TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
  
        
      
    If TmpInteger = 0 Then
         CBWTest_New_3130 = 2  'write fail
          Exit Function
     End If
    
      
    TmpInteger = Write_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_3130 = 2  'write fail
        Exit Function
    End If
    
    TmpInteger = Read_Data(LBA, Lun, CBWDataTransferLength)
     
    If TmpInteger = 0 Then
        CBWTest_New_3130 = 3    'Read fail
        Exit Function
    End If
     
    For i = 0 To CBWDataTransferLength - 1
    
        If ReadData(i) <> Pattern(i) Then
          CBWTest_New_3130 = 3    'Read fail
          Exit Function
        End If
    
    Next
    
    CBWTest_New_3130 = 1
        
    
    End Function

Public Sub AU3130_1_SlotTest()
Dim i As Integer

   If ChipName = "AU3150ALF21" Or ChipName = "AU3150ILF21" Or ChipName = "AU3150CLF21" Or ChipName = "AU3150LLF21" Then
                   ChipName = "AU3150JLF21"
              End If
           
   If Left(ChipName, 10) = "AU3150KLF2" Then
                   TmpChip = ChipName
                   ChipName = "AU3150JLF21"
    End If
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                
                
                 '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &H80)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                 CardResult = DO_WritePort(card, Channel_P1A, &H2F)   ' usb Power Off
                 Call MsecDelay(0.8)    'power on time
                 
                 If Left(TmpChip, 10) = "AU3150KLF2" Then
                   Call MsecDelay(2.5)     'power on time
                 End If
                 
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
                 
                
                
                ClosePipe
                LBA = LBA + 1
                OldLBa = LBA
                LBA = LBA + 20000 ' 20000 mp3 shift
                
               
                If Left(TmpChip, 10) = "AU3150KLF2" Then
                   rv0 = CBWTest_New_31303(0, 1, "vid_058f")
                 Else
                   rv0 = CBWTest_New_3130(0, 1, "vid_058f")
                 
                 End If
                
                LBA = OldLBa
                'Print "a1"
                Call LabelMenu(0, rv0, 1)
                ClosePipe
                 
                 
                If rv0 = 1 Then
                
                        ' =========== save MP3 pattern
                                If ChipName = "AU3130" Then
                                For i = 0 To 100
                                   MP3Data(i) = MP3Data_A(i)
                                Next
                            End If
                            
                             If ChipName = "AU3130B" Or ChipName = "AU3130CL" Then
                                For i = 0 To 100
                                   MP3Data(i) = MP3Data_B(i)
                                   MP3Data_1(i) = MP3Data_B1(i)
                                   MP3Data_2(i) = MP3Data_B2(i)
                                   MP3Data_3(i) = MP3Data_C(i)
                                   MP3Data_4(i) = MP3Data_C1(i)
                                   MP3Data_5(i) = MP3Data_C2(i)
                                  
                                Next
                            End If
                            
                               If ChipName = "AU3130BLF20" Then
                                For i = 0 To 100
                                   MP3Data_6(i) = MP3Data_BL(i)
                                   MP3Data_7(i) = MP3Data_BL1(i)
                                Next
                            End If
                            
                             If ChipName = "AU3130CLF20" Then
                                For i = 0 To 100
                                   MP3Data_8(i) = MP3Data_CL(i)
                                   
                                Next
                            End If
                            
                            '(4)============================2007.06.14 CHEYENNE CHANG
                           If ChipName = "AU3130CLF20" Then
                                For i = 0 To 100
                                   MP3Data_9(i) = MP3Data_CW1(i)
                                   
                                Next
                            End If

                            
                            If ChipName = "AU3150JLF21" Then
                                For i = 0 To 100
                                   MP3Data_10(i) = MP3Data_3150J(i)
                                   MP3Data_11(i) = MP3Data_3150J1(i)
                                Next
                            End If
  
                            If ChipName = "AU3152ALF21" Or ChipName = "AU3152ALF23" Then
                                For i = 0 To 100
                                   MP3Data_15(i) = MP3Data_3152A1(i)
                                   MP3Data_16(i) = MP3Data_3152A2(i)
                                   MP3Data_161(i) = MP3Data_3152A3(i)
                                Next
                            End If
                            
                            
                             If TmpChip = "AU3150KLF21" Then
                                For i = 0 To 100
                                   MP3Data_19(i) = MP3Data_3150KL1(i)
                                   MP3Data_20(i) = MP3Data_3150KL2(i)
                                Next
                            End If
                            
                               If TmpChip = "AU3150KLF22" Then
                                For i = 0 To 100
                                   MP3Data_21(i) = MP3Data_3150KL21(i)
                                   MP3Data_22(i) = MP3Data_3150KL22(i)
                                   
                                Next
                            End If
                            
                            
                              If TmpChip = "AU3150KLF23" Then
                                For i = 0 To 100
                                   MP3Data_21(i) = MP3Data_3150KL21(i)
                                   MP3Data_22(i) = MP3Data_3150KL22(i)
                                   MP3Data_23(i) = MP3Data_3150KL23(i)
                                Next
                            End If
                            
                            
                              If TmpChip = "AU3150KLF24" Then
                                For i = 0 To 100
                                   MP3Data_21(i) = MP3Data_3150KL21(i)
                                   MP3Data_22(i) = MP3Data_3150KL22(i)
                                   MP3Data_23(i) = MP3Data_3150KL23(i)
                                   MP3WMA_23(i) = MP3WMA_3150KL23(i)
                                   MP3WMA_231(i) = MP3WMA_3150KL231(i)
                                Next
                            End If
                          '================================
                              CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                             If CardResult <> 0 Then
                                MsgBox "Power off fail"
                                End
                             End If
                            ' mp3==================================================
                             Call MsecDelay(0.5)
                             
                             CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                               
                             Call MsecDelay(4)    'power on time and DAC initial time
                               If Left(TmpChip, 10) = "AU3150KLF2" Then  'debug
                                 Call MsecDelay(2.5)     'power on time
                              End If
                 
                             
                            ' Data initial =================
           
                            CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
         
                            If Left(ChipName, 7) = "AU3130B" Then    ' MP3 Power Off
                              Call MsecDelay(0.5)
                             
                              CardResult = DO_WritePort(card, Channel_P1A, &H8D)
                           
                            End If
                            
                            
                               If Left(ChipName, 7) = "AU3150J" Or Left(ChipName, 7) = "AU3152A" Then      ' MP3 Power Off
                           '   Call MsecDelay(0.5)
                             
                              CardResult = DO_WritePort(card, Channel_P1A, &H8D)
                              
                           '   Call MsecDelay(10)
                            End If
                            
                            If Left(ChipName, 7) = "AU3130C" Then    ' MP3 Power Off
                              Call MsecDelay(2)
                              CardResult = DO_WritePort(card, Channel_P1A, &H8D)
                            End If
                           
                              If CardResult <> 0 Then
                                MsgBox "Power on fail"
                                End
                             End If
                             
                    rv2 = AU3130Test
                  Else
                   rv2 = 4
                 End If
                 
                 Call LabelMenu(2, rv2, rv0)
                 
                  
                 ' Call MsecDelay(2)
                 
                  ' Open App.Path & "\MP3.txt" For Output As #4

                    ' For i = 0 To 100
                    '  Print #4, gnBuffer(i)
                    'Next i
                  

                    'Close #4


                
                  
                '  dma2205.Hide
                 
                  
                  
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
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
                ElseIf rv0 * rv2 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub

Public Sub AU3130_2_SlotTest()
Dim i As Integer
            If ChipName = "AU3150ALF22" Or ChipName = "AU3150CLF22" Then
                   ChipName = "AU3150ALF22"
              End If
           
           
                If PCI7248InitFinish = 0 Then
                  PCI7248Exist
                End If
                
                
                 '=========================================
                '    POWER on
                '=========================================
                 CardResult = DO_WritePort(card, Channel_P1A, &H80)
                 
                 If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 
                 Call MsecDelay(0.05)
                 CardResult = DO_WritePort(card, Channel_P1A, &H2F)   ' usb Power Off
                
                 Call MsecDelay(0.8)    'power on time
                 
                  If CardResult <> 0 Then
                    MsgBox "Power on fail"
                    End
                 End If
                 
                
                
                ClosePipe
                LBA = LBA + 1
                OldLBa = LBA
                 LBA = LBA + 20000 ' 20000 mp3 shift
                rv0 = CBWTest_New_3130(0, 1, "vid_058f")
               
                 LBA = OldLBa
               
                 Call LabelMenu(0, rv0, 1)
                ClosePipe
                 rv1 = CBWTest_New_31302(1, rv0, "vid_058f")
                 Call LabelMenu(1, rv1, rv0)
                 ClosePipe
                 
                 Tester.Print " Flash R/W test rv0="; rv0
                 
                  Tester.Print "SD R/W teswt rv1="; rv1
                 
                If rv0 * rv1 = 1 Then
                
                        ' =========== save MP3 pattern
                                If ChipName = "AU3130" Then
                                For i = 0 To 100
                                   MP3Data(i) = MP3Data_A(i)
                                Next
                            End If
                            
                             If ChipName = "AU3130B" Or ChipName = "AU3130CL" Then
                                For i = 0 To 100
                                   MP3Data(i) = MP3Data_B(i)
                                   MP3Data_1(i) = MP3Data_B1(i)
                                   MP3Data_2(i) = MP3Data_B2(i)
                                   MP3Data_3(i) = MP3Data_C(i)
                                   MP3Data_4(i) = MP3Data_C1(i)
                                   MP3Data_5(i) = MP3Data_C2(i)
                                  
                                Next
                            End If
                            
                               If ChipName = "AU3130BLF20" Then
                                For i = 0 To 100
                                   MP3Data_6(i) = MP3Data_BL(i)
                                   MP3Data_7(i) = MP3Data_BL1(i)
                                Next
                            End If
                            
                             If ChipName = "AU3130CLF20" Then
                                For i = 0 To 100
                                   MP3Data_8(i) = MP3Data_CL(i)
                                   
                                Next
                            End If
                            
                            '(4)============================2007.06.14 CHEYENNE CHANG
                           If ChipName = "AU3130CLF20" Then
                                For i = 0 To 100
                                   MP3Data_9(i) = MP3Data_CW1(i)
                                   
                                Next
                            End If

                            
                            If ChipName = "AU3150JLF21" Then
                                For i = 0 To 100
                                   MP3Data_10(i) = MP3Data_3150J(i)
                                   MP3Data_11(i) = MP3Data_3150J1(i)
                                Next
                            End If
  
                            If ChipName = "AU3152ALF21" Or ChipName = "AU3152ALF23" Then
                                For i = 0 To 100
                                   MP3Data_15(i) = MP3Data_3152A1(i)
                                   MP3Data_16(i) = MP3Data_3152A2(i)
                                Next
                            End If
                            
                            
                               If ChipName = "AU3150ALF22" Then
                                For i = 0 To 100
                                   MP3Data_17(i) = MP3Data_3150A221(i)  ' for AU3150ALF22, AU3150CLF22 , because rom code
                                   MP3Data_18(i) = MP3Data_3150A222(i)
                                Next
                            End If
                            
                          '================================
                              CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
                             If CardResult <> 0 Then
                                MsgBox "Power off fail"
                                End
                             End If
                            ' mp3==================================================
                             Call MsecDelay(0.5)
                             
                             CardResult = DO_WritePort(card, Channel_P1A, &H9F)   ' MP3 Power Off
                               
                             Call MsecDelay(4)    'power on time and DAC initial time
                             
                             
                            ' Data initial =================
           
                            CardResult = DO_WritePort(card, Channel_P1A, &H8F)
                             
         
                            If Left(ChipName, 7) = "AU3130B" Then    ' MP3 Power Off
                              Call MsecDelay(0.5)
                             
                              CardResult = DO_WritePort(card, Channel_P1A, &H8D)
                           
                            End If
                            
                            
                               If Left(ChipName, 7) = "AU3150J" Or Left(ChipName, 7) = "AU3152A" Or ChipName = "AU3152ALF22" Then          ' MP3 Power Off
                           '   Call MsecDelay(0.5)
                             
                              CardResult = DO_WritePort(card, Channel_P1A, &H8D)
                              
                           '   Call MsecDelay(10)
                            End If
                            
                            If Left(ChipName, 7) = "AU3130C" Then    ' MP3 Power Off
                              Call MsecDelay(2)
                              CardResult = DO_WritePort(card, Channel_P1A, &H8D)
                            End If
                           
                              If CardResult <> 0 Then
                                MsgBox "Power on fail"
                                End
                             End If
                             
                    rv2 = AU3130Test
                  Else
                   rv2 = 4
                 End If
                 
                 Call LabelMenu(2, rv2, rv1)
                 
                  
                  ' Call MsecDelay(2)
                 
                     'Open App.Path & "\MP3.txt" For Output As #4

                 '     For i = 0 To 100
                 '    Print #4, gnBuffer(i)
                 '    Next i
                  

               '      Close #4


                
                  
                '  dma2205.Hide
                 
                  
                  
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
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
                ElseIf rv0 * rv1 * rv2 = 1 Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
End Sub


Public Sub CompareDataOld(AudioType As Byte)
 Dim ErrCount As Integer
 Dim OldCount As Integer
 Dim CoutinueCount As Integer
 
MP3CompareResult = 1
WMACompareResult = 1

Dim k As Integer
Dim j As Integer
Dim i As Integer
Dim PatternErrSum As Double
Dim PatternErr As Single
Dim PatternDirection As Integer
Dim RawDataErrSum As Double
Dim RawDataDirection As Integer
Dim RawDataErr As Single
 
If AudioType = 0 Then
    
    For k = 0 To MP3PatternIndex - 1
    
          
           
           For j = 0 To 1     ' channel index
           
                PatternErrSum = 0
                PatternDirection = 0
                RawDataErrSum = 0
                RawDataDirection = 0
                
                
                For i = 1 To 99
          
                       
                        If Abs(MP3Pattern(k, j, i)) > 500 Then
                        
                            PatternErr = (MP3Pattern(k, j, i) - MP3Pattern(k, j, i - 1)) / MP3Pattern(k, j, i)
                            
                            
                            If PatternErr >= 0 Then
                            
                               PatternDirection = PatternDirection + 1
                            Else
                               PatternDirection = PatternDirection - 1
                            End If
                            
                            PatternErrSum = PatternErrSum + PatternErr * PatternErr
                            
                            
                            RawDataErr = (MP3RawData(j, i) - MP3RawData(j, i - 1)) / MP3Pattern(k, j, i)
                            
                            If RawDataErr >= 0 Then
                              
                               RawDataDirection = RawDataDirection + 1
                            Else
                               RawDataDirection = RawDataDirection - 1
                            End If
                            
                            RawDataErrSum = RawDataErrSum + RawDataErr * RawDataErr
                       End If
                        
                  
                Next i
                
                
                
                Tester.Print "MP3 Pattern Dir="; PatternDirection
                Tester.Print "MP3 RawData Dir="; RawDataDirection
                
                Tester.Print "MP3 Pattern Err="; PatternErrSum
                Tester.Print "MP3 RawData Err="; RawDataErrSum
                
                
                If Abs(PatternDirection - RawDataDirection) <= 2 And ((RawDataErrSum / PatternErrSum) < 1.5 Or (RawDataErrSum / PatternErrSum) > 0.5) Then
                        MP3MatchIndex = k ' get correct pattern
                        Tester.Cls
                        Tester.Print "MP3 mode="; k
                        
                        Exit Sub
                End If
           
           Next j
       
    Next k
    ' compare fail
   MP3CompareResult = 2
End If
 
 
'=====================================================================
'
'   WMA process
'
'=====================================================================




If AudioType = 1 Then
    
    For k = 0 To WMAPatternIndex - 1
    
           
           
             For j = 0 To 1     ' channel index
           
                PatternErrSum = 0
                PatternDirection = 0
                RawDataErrSum = 0
                RawDataDirection = 0
                For i = 1 To 99
          
                       
                        If Abs(WMAPattern(k, j, i)) > 500 Then
                        
                            PatternErr = (WMAPattern(k, j, i) - WMAPattern(k, j, i - 1)) / WMAPattern(k, j, i)
                            
                            
                            If PatternErr >= 0 Then
                            
                               PatternDirection = PatternDirection + 1
                            Else
                               PatternDirection = PatternDirection - 1
                            End If
                            
                            PatternErrSum = PatternErrSum + PatternErr * PatternErr
                            
                            
                            RawDataErr = (WMARawData(j, i) - WMARawData(j, i - 1)) / WMAPattern(k, j, i)
                            
                            If RawDataErr >= 0 Then
                              
                               RawDataDirection = RawDataDirection + 1
                            Else
                               RawDataDirection = RawDataDirection - 1
                            End If
                            
                            RawDataErrSum = RawDataErrSum + RawDataErr * RawDataErr
                       End If
                        
                  
                Next i
                
                
                
                Tester.Print "WMA Pattern Dir="; PatternDirection
                Tester.Print "WMA RawData Dir="; RawDataDirection
                
                Tester.Print "WMA Pattern Err="; PatternErrSum
                Tester.Print "WMA RawData Err="; RawDataErrSum
                
                
                If Abs(PatternDirection - RawDataDirection) <= 2 And ((RawDataErrSum / PatternErrSum) < 1.5 Or (RawDataErrSum / PatternErrSum) > 0.5) Then
                        WMAMatchIndex = k ' get correct pattern
                        Tester.Print "WMA mode="; k
                        Exit Sub
                End If
           
           Next j
       
    Next k
    WMACompareResult = 2
End If







End Sub
Public Sub DistingDataTraceOld(AudioType As Byte)


 On Error Resume Next
 Dim ErrCount As Integer
 Dim OldCount As Integer
 Dim ContinueCount As Integer
 
MP3DistinguishResult = 1
WMADistinguishResult = 1

Dim k As Integer
Dim j As Integer
Dim i As Integer
Dim PassCount As Integer

Dim PatternAbsSum As Double
Dim RawDataAbsSum As Double
Dim DiffSum As Double
Dim DiffAvg As Double
Dim AmpRatio As Double
Dim Diff As Long
Dim DiffFail As Byte
Dim AmpFail As Byte
Dim AvgFail As Byte
Dim Avg As Single
Dim PeakFail As Byte
 
Dim SineTop As Long
Dim SineBottom As Long
Dim SineAmp As Long
 
Dim X As Double
Dim Z As Double
Dim Delta As Double
Dim DeltaY As Long
Dim SineShape As Single
Dim Criteria As Long
Dim OldSineShape As Single
Dim DistinguishBegin As Byte
Criteria = 5000
SineTop = -65536
SineBottom = 65536

If AudioType = 0 Then
   
    For j = 0 To 1
                 
                 ' find top and bottom
                 For i = 0 To 5999
                          SineCurve(j, i) = 0
                          SineCurve2(j, i) = 0
                        
                        ' Find SineTop
                             If SineTop < MP3RawData(j, i) Then
                               SineTop = MP3RawData(j, i)  'the max high
                              
                            End If
                          
                         ' find Bottom
                        
                          If SineBottom > MP3RawData(j, i) Then
                                   SineBottom = MP3RawData(j, i) 'the min Bottom
                                   
                           End If
                         
                  Next i
                  
                  '  no data error
                  If Abs(SineTop) < 1000 Or Abs(SineBottom) < 1000 Then
                      MP3DistinguishResult = 2
                      Exit Sub
                  End If
                  ' get amplitude
                    
                  If Abs(SineTop) <= 32767 And Abs(SineBottom) <= 32767 Then
                  
                      SineAmp = (Abs(SineTop) + Abs(SineBottom)) * 0.5
                      DeltaY = SineTop - SineAmp
                  Else
                       MP3DistinguishResult = 2
                      Exit Sub
                  End If
                    
                    
                    
                  ' get Delta Y
              '    Delat = ArcSin(MP3RawData(j, 0) / SineAmp)
                  
                  'ArcSin(X) = Atn(X / Sqr(-X * X + 1))
                  
                  X = (MP3RawData(j, 0) - DeltaY) / SineAmp
                  Z = Sqr(-X * X + 1)
                  If Abs(Z) <> 0 Then
                  
                     Delta = Atn(X / Z)
                  Else
                     Delta = 3.141596 / 2#
                  End If
                  
               ' get Delta direction
                   Avg = 0
                  For i = 0 To 9
                   Avg = Avg + MP3RawData(j, i)
                  Next i
                   Avg = Avg * 0.1
                     
                     
                   If Avg < MP3RawData(j, 0) Then
                     Delta = 3.14159 - Delta
                   End If
                     
                    
                ' compare sine wave
                 
                 For i = 0 To 5999
                  '  3.14159 / 200=0.01570795
                  
                  
                  SineShape = Sin(0.01570795 * i + Delta)
                  SineCurve(j, i) = SineAmp * SineShape + DeltaY
                  
              '    If Abs(SineShape) > 0.1 Then
                  
              '       Criteria = 5000
              '    Else
                  
              '       Criteria = 3000
              '    End If
                  
                  
                  
                  If Abs(SineCurve(j, i) - MP3RawData(j, i)) > Criteria Then
                          
                           MP3DistinguishResult = 2
                           Exit Sub
                  End If
                  
                        
                        
                            
                 Next i
      
     Next j
             
              
End If
 
 
'=====================================================================
'
'   WMA process
'
'=====================================================================




If AudioType = 1 Then
    
      For j = 0 To 1
                 
                 ' find top and bottom
                 For i = 0 To 5999
                          
                         
                        
                        ' Find SineTop
                             If SineTop < WMARawData(j, i) Then
                               SineTop = WMARawData(j, i)  'the max high
                              
                            End If
                          
                         ' find Bottom
                        
                          If SineBottom > WMARawData(j, i) Then
                                   SineBottom = WMARawData(j, i) 'the min Bottom
                                   
                           End If
                         
                  Next i
                  
                  '  no data error
                  If Abs(SineTop) < 1000 Or Abs(SineBottom) < 1000 Then
                      WMADistinguishResult = 2
                      Exit Sub
                  End If
                  ' get amplitude
                    
                  If Abs(SineTop) <= 32767 And Abs(SineBottom) <= 32767 Then
                  
                      SineAmp = (Abs(SineTop) + Abs(SineBottom)) * 0.5
                      DeltaY = SineTop - SineAmp
                  Else
                     WMADistinguishResult = 2
                      Exit Sub
                      
                  End If
                  
                  
                    
                    
                    
                  ' get Delta Y
              '    Delat = ArcSin(WMARawData(j, 0) / SineAmp)
                  
                  'ArcSin(X) = Atn(X / Sqr(-X * X + 1))
                  
                  X = (WMARawData(j, 0) - DeltaY) / SineAmp
                  Z = Sqr(-X * X + 1)
                  If Abs(Z) <> 0 Then
                  
                     Delta = Atn(X / Z)
                  Else
                     Delta = 3.141596 / 2#
                  End If
                  
               ' get Delta direction
                   Avg = 0
                  For i = 0 To 9
                   Avg = Avg + WMARawData(j, i)
                  Next i
                   Avg = Avg * 0.1
                     
                     
                   If Avg < WMARawData(j, 0) Then
                     Delta = 3.14159 - Delta
                   End If
                     
                    
                ' compare sine wave
                   DistinguishBegin = 0
                 For i = 0 To 5999
                  '  3.14159 / 200=0.01570795
                  
                  
                   
               
                      SineShape = Sin(0.01570795 * i + Delta)
                      SineCurve2(j, i) = SineAmp * SineShape + DeltaY
                      
                    '    If Abs(SineShape) > 0.7 Then
                      
                    '     Criteria = 5000
                    '  Else
                      
                    '     Criteria = 3000
                    '  End If
                      
                      
                      If Abs(SineCurve2(j, i) - WMARawData(j, i)) > Criteria Then
                        
                         WMADistinguishResult = 2
                      Exit Sub
                      End If
                      
                   
                            
                 Next i
      
     Next j
    
     
End If
 
Exit Sub
Err1:
MsgBox "Data compare Err"
End Sub


Public Sub DistingDataTrace(AudioType As Byte)
 On Error Resume Next
 Dim ErrCount As Integer
 Dim OldCount As Integer
 Dim ContinueCount As Integer
 
MP3DistinguishResult = 1
WMADistinguishResult = 1

Dim k As Integer
Dim j As Integer
Dim i As Integer
Dim PassCount As Integer

Dim PatternAbsSum As Double
Dim RawDataAbsSum As Double
Dim DiffSum As Double
Dim DiffAvg As Double
Dim AmpRatio As Double
Dim Diff As Long
Dim DiffFail As Byte
Dim AmpFail As Byte
Dim AvgFail As Byte
Dim Avg As Single
Dim PeakFail As Byte
 
Dim SineTop As Long
Dim SineBottom As Long
Dim SineAmp As Long
 
Dim X As Double
Dim Z As Double
Dim Delta As Double
Dim DeltaY As Long
Dim SineShape As Single
Dim Criteria As Long
Dim OldSineShape As Single
Dim DistinguishBegin As Byte
Dim CrossPoint As Long
Dim CrossPointX(0 To 5999) As Integer
Dim SineTopX As Long
Dim SineBottomX As Long
Dim SineHead As Long
Dim SineTail As Long
Dim ki As Long
Dim BeginFindFlag As Byte
Dim ii As Long
Dim FirstSineHead As Long
Dim LastSineTail As Long
Dim CrossPointCounter As Long


Criteria = 6000
SineTop = -65536
SineBottom = 65536

If AudioType = 0 Then
   
    For j = 0 To 1
    
                 
                   For i = 0 To 5999
                   
                     CrossPointX(i) = 0
                     SineCurve(j, i) = 0
                     SineCurve2(j, i) = 0
                  Next i
                  
                  CrossPointCounter = 0
                  For i = 0 To 5990
                         
                        '======================================
                        ' find cross point
                        '======================================
                        If MP3RawData(j, i) < 0 And MP3RawData(j, i + 1) > 0 Then
                                CrossPointX(i) = 1
                                CrossPointX(i + 1) = 2
                                 Debug.Print i, CrossPointX(i)
                                 CrossPointCounter = CrossPointCounter + 1
                        End If
                    
                           
                    
                   Next i
                   
                   
                   If CrossPointCounter = 0 Then
                    MP3DistinguishResult = 2
                                 Exit Sub
                   End If
                    
                    FirstSineHead = 5999
                    LastSineTail = 0
                 
                 ' find top and bottom
                 For i = 0 To 5999
                          
                        '================================
                        ' Begin Find
                        '================================
                          
                       
                        If CrossPointX(i) = 2 Then
                            BeginFindFlag = 1
                            SineTop = -65536
                            SineBottom = 65536
                            SineHead = i
                            
                            
                            If FirstSineHead > SineHead Then
                                FirstSineHead = SineHead
                            End If
                        End If
                        
                        
                        If BeginFindFlag = 1 And CrossPointX(i) = 0 Then
                        
                          ' Find SineTop
                             If SineTop < MP3RawData(j, i) Then
                               SineTop = MP3RawData(j, i)  'the max high
                               SineTopX = i
                            End If
                          
                         ' find Bottom
                        
                          If SineBottom > MP3RawData(j, i) Then
                                SineBottom = MP3RawData(j, i) 'the min Bottom
                                SineBottomX = i
                           End If
                        End If
                        
                        
                        '=====================================
                        '  End Find
                        '====================================
                        If CrossPointX(i) = 1 And BeginFindFlag = 1 Then
                            
                            SineTail = i
                            
                             
                            If LastSineTail < SineTail Then
                                LastSineTail = SineTail
                            End If
                            '=================================
                            '  Check Amplitude
                            '=================================
                            '  no data error
                            If Abs(SineTop) < 1000 Or Abs(SineBottom) < 1000 Then
                                MP3DistinguishResult = 2
                                 Exit Sub
                            End If
                            ' get amplitude
                              
                            If Abs(SineTop) <= 32767 And Abs(SineBottom) <= 32767 Then
                            
                                SineAmp = (Abs(SineTop) + Abs(SineBottom)) * 0.5
                                DeltaY = SineTop - SineAmp
                            Else
                                 MP3DistinguishResult = 2
                                 Exit Sub
                            End If
                
                            
        
                              
                              For ii = SineHead To SineTail
                             
                              ki = SineTopX - ii
                              SineShape = Sin(0.01570795 * ki + 3.141596 / 2)
                          
                              SineCurve(j, ii) = SineAmp * SineShape + DeltaY
                            Next ii
                       End If
                       
                            
                            '==============================
                            ' make sine curve
                            '==============================
                  
                         
                  Next i
                  
               
                    
                    
            
                 For i = FirstSineHead To LastSineTail
   
                   If Abs(SineCurve(j, i) - MP3RawData(j, i)) > Criteria Then
                          
                            MP3DistinguishResult = 2
                           Exit Sub
                   End If
                
               Next i
      
     Next j
             
              
End If
 
 
'=====================================================================
'
'   WMA process
'
'=====================================================================




If AudioType = 1 Then
    
      For j = 0 To 1
      
                    
    
                 
                   For i = 0 To 5999
                   
                     CrossPointX(i) = 0
                    ' SineCurve(j, i) = 0
                    ' SineCurve2(j, i) = 0
                  Next i
                  CrossPointCounter = 0
                  
                  For i = 0 To 5990
                         
                        '======================================
                        ' find cross point
                        '======================================
                        If WMARawData(j, i) < 0 And WMARawData(j, i + 1) > 0 Then
                                CrossPointX(i) = 1
                                CrossPointX(i + 1) = 2
                                CrossPointCounter = CrossPointCounter + 1
                                 Debug.Print i, CrossPointX(i)
                                 
                        End If
                    
                           
                    
                   Next i
                 
                    If CrossPointCounter = 0 Then
                    
                     WMADistinguishResult = 2
                                 Exit Sub
                    End If
                    FirstSineHead = 5999
                    LastSineTail = 0
                 
                 ' find top and bottom
                 For i = 0 To 5999
                          
                        '================================
                        ' Begin Find
                        '================================
                          
                       
                        If CrossPointX(i) = 2 Then
                            BeginFindFlag = 1
                            SineTop = -65536
                            SineBottom = 65536
                            SineHead = i
                            
                            
                            If FirstSineHead > SineHead Then
                                FirstSineHead = SineHead
                            End If
                        End If
                        
                        
                        If BeginFindFlag = 1 And CrossPointX(i) = 0 Then
                        
                          ' Find SineTop
                             If SineTop < WMARawData(j, i) Then
                               SineTop = WMARawData(j, i)  'the max high
                               SineTopX = i
                            End If
                          
                         ' find Bottom
                        
                          If SineBottom > WMARawData(j, i) Then
                                SineBottom = WMARawData(j, i) 'the min Bottom
                                SineBottomX = i
                           End If
                        End If
                        
                        
                        '=====================================
                        '  End Find
                        '====================================
                        If CrossPointX(i) = 1 And BeginFindFlag = 1 Then
                            
                            SineTail = i
                            
                             
                            If LastSineTail < SineTail Then
                                LastSineTail = SineTail
                            End If
                            '=================================
                            '  Check Amplitude
                            '=================================
                            '  no data error
                            If Abs(SineTop) < 1000 Or Abs(SineBottom) < 1000 Then
                                WMADistinguishResult = 2
                                 Exit Sub
                            End If
                            ' get amplitude
                              
                            If Abs(SineTop) <= 32767 And Abs(SineBottom) <= 32767 Then
                            
                                SineAmp = (Abs(SineTop) + Abs(SineBottom)) * 0.5
                                DeltaY = SineTop - SineAmp
                            Else
                                 WMADistinguishResult = 2
                                 Exit Sub
                            End If
                
                            
        
                              
                              For ii = SineHead To SineTail
                             
                              ki = SineTopX - ii
                              SineShape = Sin(0.01570795 * ki + 3.141596 / 2)
                          
                              SineCurve2(j, ii) = SineAmp * SineShape + DeltaY
                            Next ii
                       End If
                       
                            
                            '==============================
                            ' make sine curve
                            '==============================
                  
                         
                  Next i
                  
               
                    
                     
            
                 For i = FirstSineHead To LastSineTail
   
                   If Abs(SineCurve2(j, i) - WMARawData(j, i)) > Criteria Then
                          
                            WMADistinguishResult = 2
                           Exit Sub
                   End If
                
               Next i
      
     Next j
    
     
End If
 
Exit Sub
Err1:
MsgBox "Data compare Err"


 





End Sub

Public Sub DistingDataTraceOld2(AudioType As Byte)
 On Error Resume Next
 Dim ErrCount As Integer
 Dim OldCount As Integer
 Dim ContinueCount As Integer
 
MP3DistinguishResult = 1
WMADistinguishResult = 1

Dim k As Integer
Dim j As Integer
Dim i As Integer
Dim PassCount As Integer

Dim PatternAbsSum As Double
Dim RawDataAbsSum As Double
Dim DiffSum As Double
Dim DiffAvg As Double
Dim AmpRatio As Double
Dim Diff As Long
Dim DiffFail As Byte
Dim AmpFail As Byte
Dim AvgFail As Byte
Dim Avg As Single
Dim PeakFail As Byte
 
Dim SineTop As Long
Dim SineBottom As Long
Dim SineAmp As Long
 
Dim X As Double
Dim Z As Double
Dim Delta As Double
Dim DeltaY As Long
Dim SineShape As Single
Dim Criteria As Long
Dim OldSineShape As Single
Dim DistinguishBegin As Byte
Dim CrossPoint As Long
Criteria = 7500
SineTop = -65536
SineBottom = 65536

If AudioType = 0 Then
   
    For j = 0 To 1
    
                 
                   For i = 0 To 5999
                          
                         
                        If MP3RawData(j, i) > 0 Then
                        '  If WMARawData(j, i + 1) < 0 Then
                        
                        '     CrossPoint = i + 1
                        '    Exit For
                        '  End If
                     
                        
                        ElseIf MP3RawData(j, i) < 0 Then
                          If MP3RawData(j, i + 1) > 0 Then
                        
                          CrossPoint = i + 1 + 40
                           Exit For
                          End If
                        Else
                        
                        End If
                        
                         
                  Next i
                 
    
                 
                 ' find top and bottom
                 For i = CrossPoint To 5999
                          SineCurve(j, i) = 0
                          SineCurve2(j, i) = 0
                        
                        ' Find SineTop
                             If SineTop < MP3RawData(j, i) Then
                               SineTop = MP3RawData(j, i)  'the max high
                              
                            End If
                          
                         ' find Bottom
                        
                          If SineBottom > MP3RawData(j, i) Then
                                   SineBottom = MP3RawData(j, i) 'the min Bottom
                                   
                           End If
                         
                  Next i
                  
                  '  no data error
                  If Abs(SineTop) < 1000 Or Abs(SineBottom) < 1000 Then
                      MP3DistinguishResult = 2
                      Exit Sub
                  End If
                  ' get amplitude
                    
                  If Abs(SineTop) <= 32767 And Abs(SineBottom) <= 32767 Then
                  
                      SineAmp = (Abs(SineTop) + Abs(SineBottom)) * 0.5
                      DeltaY = SineTop - SineAmp
                  Else
                       MP3DistinguishResult = 2
                      Exit Sub
                  End If
                    
                    
                    
                  ' get Delta Y
              '    Delat = ArcSin(MP3RawData(j, 0) / SineAmp)
                  
                  'ArcSin(X) = Atn(X / Sqr(-X * X + 1))
                  
                  X = (MP3RawData(j, CrossPoint) - DeltaY) / SineAmp
                  Z = Sqr(-X * X + 1)
                  If Abs(Z) <> 0 Then
                  
                     Delta = Atn(X / Z)
                  Else
                     Delta = 3.141596 / 2#
                  End If
                  
               ' get Delta direction
                   Avg = 0
                  For i = CrossPoint To 9 + CrossPoint
                   Avg = Avg + MP3RawData(j, i)
                  Next i
                   Avg = Avg * 0.1
                     
                     
                   If Avg < MP3RawData(j, CrossPoint) Then
                     Delta = 3.14159 - Delta
                   End If
                     
                    
                ' compare sine wave
                 
                 For i = CrossPoint To 5999
                  '  3.14159 / 200=0.01570795
                  
                  
                  SineShape = Sin(0.01570795 * i + Delta)
                  SineCurve(j, i) = SineAmp * SineShape + DeltaY
                  
              '    If Abs(SineShape) > 0.1 Then
                  
              '       Criteria = 5000
              '    Else
                  
              '       Criteria = 3000
              '    End If
                  
                  
                  
                  If Abs(SineCurve(j, i) - MP3RawData(j, i)) > Criteria Then
                          
                           MP3DistinguishResult = 2
                          Exit Sub
                  End If
                  
                        
                        
                            
                 Next i
      
     Next j
             
              
End If
 
 
'=====================================================================
'
'   WMA process
'
'=====================================================================




If AudioType = 1 Then
    
      For j = 0 To 1
      
                   For i = 0 To 5999
                          
                         
                        If WMARawData(j, i) > 0 Then
                        '  If WMARawData(j, i + 1) < 0 Then
                        
                        '     CrossPoint = i + 1
                        '    Exit For
                        '  End If
                     
                        
                        ElseIf WMARawData(j, i) < 0 Then
                          If WMARawData(j, i + 1) > 0 Then
                        
                          CrossPoint = i + 1 + 40
                           Exit For
                          End If
                        Else
                        
                        End If
                        
                         
                  Next i
                 
                 ' find top and bottom
                 For i = CrossPoint To 5999
                          
                         
                        
                        ' Find SineTop
                             If SineTop < WMARawData(j, i) Then
                               SineTop = WMARawData(j, i)  'the max high
                              
                            End If
                          
                         ' find Bottom
                        
                          If SineBottom > WMARawData(j, i) Then
                                   SineBottom = WMARawData(j, i) 'the min Bottom
                                   
                           End If
                         
                  Next i
                  
                  '  no data error
                  If Abs(SineTop) < 1000 Or Abs(SineBottom) < 1000 Then
                      WMADistinguishResult = 2
                      Exit Sub
                  End If
                  ' get amplitude
                    
                  If Abs(SineTop) <= 32767 And Abs(SineBottom) <= 32767 Then
                  
                      SineAmp = (Abs(SineTop) + Abs(SineBottom)) * 0.5
                      DeltaY = SineTop - SineAmp
                  Else
                     WMADistinguishResult = 2
                      Exit Sub
                      
                  End If
                  
                  
                    
                    
                    
                  ' get Delta Y
              '    Delat = ArcSin(WMARawData(j, 0) / SineAmp)
                  
                  'ArcSin(X) = Atn(X / Sqr(-X * X + 1))
                  
                  X = (WMARawData(j, CrossPoint) - DeltaY) / SineAmp
                  Z = Sqr(-X * X + 1)
                  If Abs(Z) <> 0 Then
                  
                     Delta = Atn(X / Z)
                  Else
                     Delta = 3.141596 / 2#
                  End If
                  
               ' get Delta direction
                   Avg = 0
                  For i = CrossPoint To CrossPoint + 9
                   Avg = Avg + WMARawData(j, i)
                  Next i
                   Avg = Avg * 0.1
                     
                     
                     If Avg < WMARawData(j, CrossPoint) Then
                       Delta = 3.14159 - Delta
                    End If
                     
                  
                ' compare sine wave
                    
                 For i = CrossPoint To 5999
                  '  3.14159 / 200=0.01570795
                  
                  
                   
               
                      SineShape = Sin(0.01570795 * i + Delta)
                      SineCurve2(j, i) = SineAmp * SineShape + DeltaY
                      
                    '    If Abs(SineShape) > 0.7 Then
                      
                    '     Criteria = 5000
                    '  Else
                      
                    '     Criteria = 3000
                    '  End If
                      
                      
                      If Abs(SineCurve2(j, i) - WMARawData(j, i)) > Criteria Then
                        
                         WMADistinguishResult = 2
                      Exit Sub
                      End If
                      
                   
                            
                 Next i
      
     Next j
    
     
End If
 
Exit Sub
Err1:
MsgBox "Data compare Err"


 





End Sub
Public Sub CompareDataTrace(AudioType As Byte)
On Error Resume Next
 Dim ErrCount As Integer
 Dim OldCount As Integer
 Dim ContinueCount As Integer
 
MP3CompareResult = 1
WMACompareResult = 1

Dim k As Integer
Dim j As Integer
Dim i As Integer
Dim PassCount As Integer

Dim PatternAbsSum As Double
Dim RawDataAbsSum As Double
Dim DiffSum As Double
Dim DiffAvg As Double
Dim AmpRatio As Double
Dim Diff As Long
Dim DiffFail As Byte
Dim AmpFail As Byte
Dim AvgFail As Byte
Dim Avg As Single
Dim PeakFail As Byte
 


If AudioType = 0 Then
    
    For k = 0 To MP3PatternIndex - 1
    
              
           
           For j = 0 To 1
                 
                  ' iniital
                  ErrCount = 0
                  ContinueCount = 0
                  OldCount = 0
                  AvgFail = 0
                  PeakFail = 0
                  DiffFail = 0
                 For i = 0 To 5999
                 
                 
                          
                            Diff = Abs(MP3Pattern(k, j, i) - MP3RawData(j, i))
                            If Diff > 2000 Then
                                         DiffFail = 1
                             End If
                           
                            
                    '           If MP3RawData(j, i) > MP3RawData(j, i - 1) Then
                    '            If MP3RawData(j, i) > MP3RawData(j, i + 1) Then
                    '             Avg = (CSng(MP3RawData(j, i + 1)) + CSng(MP3RawData(j, i - 1))) * 0.5
                    '              AmpRatio = CSng(MP3RawData(j, i)) / Avg
                    '                  If AmpRatio < 0.85 Or AmpRatio > 1.15 Then
                    '                   PeakFail = 1
                    '                   Exit For
                    '                  End If
                    '            End If
                    '         End If
                               
                    '         If MP3RawData(j, i) < MP3RawData(j, i - 1) Then
                    '           If MP3RawData(j, i) < MP3RawData(j, i + 1) Then
                    '              Avg = (CSng(MP3RawData(j, i + 1)) + CSng(MP3RawData(j, i - 1))) * 0.5
                    '              AmpRatio = CSng(MP3RawData(j, i)) / Avg
                    '                  If AmpRatio < 0.85 Or AmpRatio > 1.15 Then
                    '                   PeakFail = 2
                    '                   Exit For
                    '                  End If
                     '           End If
                     '        End If
                               
                            
                           
                          
                            
                 Next i
                
                
                    Tester.Print j; " DiffFail="; DiffFail
                    Tester.Print j; " PeakFail="; PeakFail
             
                
                    
          
                If DiffFail = 0 And PeakFail = 0 Then
              
                      If j = 1 Then
                        MP3MatchIndex = k ' get correct pattern
                        Tester.Print "MP3 mode="; k
                        Exit Sub
                      End If
                Else
                
                        Exit For
                End If
                    
           Next j
           
       
    Next k
   
   MP3CompareResult = 2
End If
 
 
'=====================================================================
'
'   WMA process
'
'=====================================================================




If AudioType = 1 Then
    
    For k = 0 To WMAPatternIndex - 1
    
            For j = 0 To 1
                 
                  ' iniital
                  ErrCount = 0
                  ContinueCount = 0
                  OldCount = 0
                  PeakFail = 0
                   AvgFail = 0
                   DiffFail = 0
                 For i = 0 To 5999
                 
                 
                             Diff = Abs(WMAPattern(k, j, i) - WMARawData(j, i))
                            If Diff > 2000 Then
                                         DiffFail = 1
                             End If
                          
                             
                     '        If WMARawData(j, i) > WMARawData(j, i - 1) Then
                     '          If WMARawData(j, i) > WMARawData(j, i + 1) Then
                     '            Avg = (CSng(WMARawData(j, i + 1)) + CSng(WMARawData(j, i - 1))) * 0.5
                     '             AmpRatio = CSng(WMARawData(j, i)) / Avg
                     '                If AmpRatio < 0.85 Or AmpRatio > 1.15 Then
                     '                  PeakFail = 1
                     '                 Exit For
                     '                End If
                     '           End If
                     '        End If
                               
                     '         If WMARawData(j, i) < WMARawData(j, i - 1) Then
                     '          If WMARawData(j, i) < WMARawData(j, i + 1) Then
                     '             Avg = (CSng(WMARawData(j, i + 1)) + CSng(WMARawData(j, i - 1))) * 0.5
                     '             AmpRatio = CSng(WMARawData(j, i)) / Avg
                     '                 If AmpRatio < 0.85 Or AmpRatio > 1.15 Then
                     '                  PeakFail = 2
                     '''                  Exit For
                     '                 End If
                     '           End If
                     '        End If
                         
                 Next i
                
                
                    
                     Tester.Print j; " DiffFail="; DiffFail
                    Tester.Print j; " PeakFail="; PeakFail
                
                    
          
                If DiffFail = 0 And PeakFail = 0 Then
              
                       If j = 1 Then  ' the 2nd channel
                            WMAMatchIndex = k ' get correct pattern
                            Tester.Print "WMA mode="; k
                       
                        Exit Sub
                       End If
                Else
                        Exit For
                End If
                    
           Next j
           
           
       
    Next k
    ' compare fail
   WMACompareResult = 2
End If
 
Exit Sub
Err1:
MsgBox "Data compare Err"





 





End Sub
Public Sub CompareData0919(AudioType As Byte)
On Error Resume Next
 Dim ErrCount As Integer
 Dim OldCount As Integer
 Dim ContinueCount As Integer
 
MP3CompareResult = 1
WMACompareResult = 1

Dim k As Integer
Dim j As Integer
Dim i As Integer
Dim PassCount As Integer

Dim PatternAbsSum As Double
Dim RawDataAbsSum As Double
Dim DiffSum As Double
Dim DiffAvg As Double
Dim AmpRatio As Double
Dim Diff As Single
Dim DiffFail As Byte
Dim AmpFail As Byte
Dim AvgFail As Byte
Dim Avg As Single
Dim PeakFail As Byte


If AudioType = 0 Then
    
    For k = 0 To MP3PatternIndex - 1
    
              
           
           For j = 0 To 1
                 
                  ' iniital
                  ErrCount = 0
                  ContinueCount = 0
                  OldCount = 0
                  AvgFail = 0
                  PeakFail = 0
                 For i = 1 To 5998
                 
                 
                           If CSng(MP3RawData(j, i)) > 4000 Then
                           
                                AmpRatio = CSng(MP3Pattern(k, j, i)) / CSng(MP3RawData(j, i))
                                If AmpRatio < 0.8 Or AmpRatio > 1.2 Then
                                         ErrCount = ErrCount + 1
                                         If i = OldCount + 1 Then
                                             ContinueCount = ContinueCount + 1
                                             OldCount = i
                                         End If
                                 End If
                           
                            
                            If MP3RawData(j, i) > MP3RawData(j, i - 1) Then
                              If MP3RawData(j, i) > MP3RawData(j, i + 1) Then
                                Avg = (CSng(MP3RawData(j, i + 1)) + CSng(MP3RawData(j, i - 1))) * 0.5
                                 AmpRatio = CSng(MP3RawData(j, i)) / Avg
                                     If AmpRatio < 0.85 Or AmpRatio > 1.15 Then
                                      PeakFail = 1
                                      Exit For
                                     End If
                               End If
                            End If
                               
                             If MP3RawData(j, i) < MP3RawData(j, i - 1) Then
                              If MP3RawData(j, i) < MP3RawData(j, i + 1) Then
                                 Avg = (CSng(MP3RawData(j, i + 1)) + CSng(MP3RawData(j, i - 1))) * 0.5
                                 AmpRatio = CSng(MP3RawData(j, i)) / Avg
                                     If AmpRatio < 0.85 Or AmpRatio > 1.15 Then
                                      PeakFail = 1
                                      Exit For
                                     End If
                               End If
                            End If
                               
                            
                          End If
                          
                            
                 Next i
                
                
                    Tester.Print j; " ErrCount="; ErrCount
                    Tester.Print j; " ContinueCount="; ContinueCount
             
                
                    
          
                If ErrCount < 30 And ContinueCount < 2 And PeakFail = 0 Then
              
                      If j = 1 Then
                        MP3MatchIndex = k ' get correct pattern
                        Tester.Print "MP3 mode="; k
                        Exit Sub
                      End If
                Else
                
                        Exit For
                End If
                    
           Next j
           
       
    Next k
   
   MP3CompareResult = 2
End If
 
 
'=====================================================================
'
'   WMA process
'
'=====================================================================




If AudioType = 1 Then
    
    For k = 0 To WMAPatternIndex - 1
    
            For j = 0 To 1
                 
                  ' iniital
                  ErrCount = 0
                  ContinueCount = 0
                  OldCount = 0
                  PeakFail = 0
                   AvgFail = 0
                 For i = 1 To 5998
                 
                 
                           If CSng(WMARawData(j, i)) > 4000 Then
                           
                                AmpRatio = CSng(WMAPattern(k, j, i)) / CSng(WMARawData(j, i))
                                If AmpRatio < 0.8 Or AmpRatio > 1.2 Then
                                         ErrCount = ErrCount + 1
                                         If i = OldCount + 1 Then
                                             ContinueCount = ContinueCount + 1
                                             OldCount = i
                                         End If
                                 End If
                          
                             
                            If WMARawData(j, i) > WMARawData(j, i - 1) Then
                              If WMARawData(j, i) > WMARawData(j, i + 1) Then
                                Avg = (CSng(WMARawData(j, i + 1)) + CSng(WMARawData(j, i - 1))) * 0.5
                                 AmpRatio = CSng(WMARawData(j, i)) / Avg
                                     If AmpRatio < 0.85 Or AmpRatio > 1.15 Then
                                      PeakFail = 1
                                      Exit For
                                     End If
                               End If
                            End If
                               
                             If WMARawData(j, i) < WMARawData(j, i - 1) Then
                              If WMARawData(j, i) < WMARawData(j, i + 1) Then
                                 Avg = (CSng(WMARawData(j, i + 1)) + CSng(WMARawData(j, i - 1))) * 0.5
                                 AmpRatio = CSng(WMARawData(j, i)) / Avg
                                     If AmpRatio < 0.85 Or AmpRatio > 1.15 Then
                                      PeakFail = 1
                                      Exit For
                                     End If
                               End If
                            End If
                        End If
                 Next i
                
                
                    Tester.Print j; " ErrCount="; ErrCount
                    Tester.Print j; " ContinueCount="; ContinueCount
             
                
                    
          
                If ErrCount < 30 And ContinueCount < 2 And AvgFail = 0 Then
              
                       If j = 1 Then  ' the 2nd channel
                            WMAMatchIndex = k ' get correct pattern
                            Tester.Print "WMA mode="; k
                       
                        Exit Sub
                       End If
                Else
                        Exit For
                End If
                    
           Next j
           
           
       
    Next k
    ' compare fail
   WMACompareResult = 2
End If
 
Exit Sub
Err1:
MsgBox "Data compare Err"





 





End Sub
Public Sub CompareData(AudioType As Byte)
On Error Resume Next
 Dim ErrCount As Integer
 Dim OldCount As Integer
 Dim CoutinueCount As Integer
 
MP3CompareResult = 1
WMACompareResult = 1

Dim k As Integer
Dim j As Integer
Dim i As Integer

Dim PatternAbsSum As Double
Dim RawDataAbsSum As Double
Dim DiffSum As Double
Dim DiffAvg As Double
Dim AmpRatio As Double
Dim Diff(0 To 1999) As Long
Dim DiffFail As Byte
Dim AmpFail As Byte


If AudioType = 0 Then
    
    For k = 0 To MP3PatternIndex - 1
    
              
           
           For j = 0 To 1
                 
                  ' iniital
                  PatternAbsSum = 0
                  RawDataAbsSum = 0
                  DiffSum = 0
                  DiffFail = 0
                  AmpFail = 0
         
                 For i = 0 To 5999
                          PatternAbsSum = PatternAbsSum + Abs(MP3Pattern(k, j, i))
                          RawDataAbsSum = RawDataAbsSum + Abs(MP3RawData(j, i))
                          Diff(i) = Abs(MP3Pattern(k, j, i) - MP3RawData(j, i))
                          DiffSum = DiffSum + Diff(i)
                 Next i
                
                 DiffAvg = (DiffSum / 2000)
                 AmpRatio = PatternAbsSum / RawDataAbsSum
      
      
               ' For i = 1 To 1999
               '   If 4000 < Diff(i) Then
               '     DiffFail = 1
               '   Exit For
               '   End If
               ' Next
                
                If DiffFail = 0 Then
                   If (AmpRatio > 1.01) Or (AmpRatio < 0.99) Then
                       AmpFail = 1
                   End If
                End If
                    
             '   Tester.Print "MP3 PatternAbsSum  ="; PatternAbsSum
              '  Tester.Print "MP3 RawDataAbsSum  ="; RawDataAbsSum
               ' Tester.Print "MP3 DiffAvg  ="; DiffAvg
            
                Tester.Print j; "MP3 AmpRatio  ="; AmpRatio
               
                If DiffFail = 0 And AmpFail = 0 Then
              
                      If j = 1 Then
                        MP3MatchIndex = k ' get correct pattern
                        Tester.Print "MP3 mode="; k
                        Exit Sub
                      End If
                Else
                
                        Exit For
                End If
                    
           Next j
           
       
    Next k
   
   MP3CompareResult = 2
End If
 
 
'=====================================================================
'
'   WMA process
'
'=====================================================================




If AudioType = 1 Then
    
    For k = 0 To WMAPatternIndex - 1
    
            For j = 0 To 1
                 
                  ' iniital
                  PatternAbsSum = 0
                  RawDataAbsSum = 0
                  DiffSum = 0
                  DiffFail = 0
                  AmpFail = 0
         
                 For i = 0 To 5999
                          PatternAbsSum = PatternAbsSum + Abs(WMAPattern(k, j, i))
                          RawDataAbsSum = RawDataAbsSum + Abs(WMARawData(j, i))
                          Diff(i) = Abs(WMAPattern(k, j, i) - WMARawData(j, i))
                          DiffSum = DiffSum + Diff(i)
                 Next i
                
                 DiffAvg = (DiffSum / 2000)
                 AmpRatio = PatternAbsSum / RawDataAbsSum
      
      
               ' For i = 1 To 1999
               '   If 4000 < Diff(i) Then
               '     DiffFail = 1
               '   Exit For
               '   End If
               ' Next
                
                If DiffFail = 0 Then
                   If (AmpRatio > 1.01) Or (AmpRatio < 0.99) Then
                       AmpFail = 1
                   End If
                End If
                    
        '        Tester.Print "WMA PatternAbsSum  ="; PatternAbsSum
        '        Tester.Print "WMA RawDataAbsSum  ="; RawDataAbsSum
              '  Tester.Print "WMA DiffAvg  ="; DiffAvg
                 Tester.Print j; " WMA AmpRatio  ="; AmpRatio
                
               
                If DiffFail = 0 And AmpFail = 0 Then
                    
                       If j = 1 Then  ' the 2nd channel
                            WMAMatchIndex = k ' get correct pattern
                            Tester.Print "WMA mode="; k
                       
                        Exit Sub
                       End If
                Else
                        Exit For
                End If
                    
           Next j
           
           
       
    Next k
    ' compare fail
   WMACompareResult = 2
End If
 
Exit Sub
Err1:
MsgBox "Data compare Err"





 





End Sub
 

Public Sub CompareDataOld3(AudioType As Byte)
On Error Resume Next
 Dim ErrCount As Integer
 Dim OldCount As Integer
 Dim CoutinueCount As Integer
 
MP3CompareResult = 1
WMACompareResult = 1

Dim k As Integer
Dim j As Integer
Dim i As Integer
Dim PatternErrSum As Double
Dim PatternErr As Single
Dim PatternDirection As Integer
Dim RawDataErrSum As Double
Dim RawDataDirection As Integer
Dim RawDataErr As Single
Dim MP3LErr As Single
Dim MP3RErr As Single
Dim DirLErr As Integer
Dim DirRErr As Integer
 
Dim WMALErr As Single
Dim WMARErr As Single
Dim FailFlag As Byte
If AudioType = 0 Then
    
    For k = 0 To MP3PatternIndex - 1
    
              FailFlag = 0
           
           For j = 0 To 1     ' channel index
                 
                  PatternDirection = 0
                   RawDataDirection = 0
                
                If Abs(MP3Pattern(k, j, 0)) > 5000 Then
                  
                 RawDataErr = ((CSng(MP3RawData(j, 0)) - CSng(MP3Pattern(k, j, 0)))) / CSng(MP3Pattern(k, j, 0))
                 RawDataErr = Abs(RawDataErr)
                 RawDataErrSum = RawDataErr
                  If RawDataErr > 0.5 Then
                             FailFlag = 1
                  End If
                Else
                 
                   
                 RawDataErrSum = 0
                End If
                
                
                For i = 1 To 199
                         ' Debug.Print MP3RawData(j, i)
                       
                        If Abs(MP3Pattern(k, j, i)) > 0 Then
                          
                     
                             RawDataErr = ((CSng(MP3RawData(j, i)) - CSng(MP3Pattern(k, j, i))) / CSng(MP3Pattern(k, j, i)))
                             
                             RawDataErr = Abs(RawDataErr)
                           
                             
                            If RawDataErr > 0.5 Then
                              FailFlag = 1
                      '          Debug.Print RawDataErr, MP3RawData(j, i), MP3Pattern(k, j, i)
                            End If
                              
                              RawDataErrSum = RawDataErrSum + RawDataErr
                             
                             
                            ' direction
                            If (MP3Pattern(k, j, i) - MP3Pattern(k, j, i - 1)) >= 0 Then
                            
                               PatternDirection = PatternDirection + 1
                            Else
                               PatternDirection = PatternDirection - 1
                            End If
                            
                           
                            
                            If (MP3RawData(j, i) - MP3RawData(j, i - 1)) >= 0 Then
                              
                               RawDataDirection = RawDataDirection + 1
                            Else
                               RawDataDirection = RawDataDirection - 1
                            End If
                           
                       End If
                        
                  
                Next i
                
      
              '  Tester.Print "MP3 Pattern Dir="; PatternDirection
               ' Tester.Print "MP3 RawData Dir="; RawDataDirection
                
                Tester.Print "MP3 RawData Err="; RawDataErrSum
                
               
                
                If j = 0 Then
                        MP3LErr = RawDataErrSum
                        DirLErr = Abs(PatternDirection - RawDataDirection)
                Else
                       MP3RErr = RawDataErrSum
                       DirRErr = Abs(PatternDirection - RawDataDirection)
                End If
                
                   
                If (MP3LErr < 10) And (MP3RErr < 10) And j = 1 Then
             '   If DirLErr <= 2 And DirRErr <= 2 And flag = 0 Then
                        MP3MatchIndex = k ' get correct pattern
                      '  Tester.Cls
                        Tester.Print "MP3 mode="; k
                        
                        Exit Sub
              '  End If
                End If
                    
           Next j
           
       
    Next k
    ' compare fail
   MP3CompareResult = 2
End If
 
 
'=====================================================================
'
'   WMA process
'
'=====================================================================




If AudioType = 1 Then
    
    For k = 0 To WMAPatternIndex - 1
    
              FailFlag = 0
           
           For j = 0 To 1     ' channel index
                 
                  PatternDirection = 0
                   RawDataDirection = 0
                
                If Abs(WMAPattern(k, j, 0)) > 0 Then
                  
                 RawDataErr = ((WMARawData(j, 0) - WMAPattern(k, j, 0))) / WMAPattern(k, j, 0)
                 RawDataErr = Abs(RawDataErr)
                 RawDataErrSum = RawDataErr
                  If RawDataErr > 0.5 Then
                             FailFlag = 1
                  End If
                Else
                 
                   
                 RawDataErrSum = 0
                End If
                
                
                For i = 1 To 199
          
                       
                        If Abs(WMAPattern(k, j, i)) > 0 Then
                     
                             RawDataErr = ((WMARawData(j, i) - WMAPattern(k, j, i)) / WMAPattern(k, j, i))
                             
                             RawDataErr = Abs(RawDataErr)
                           
                             
                            If RawDataErr > 0.5 Then
                              FailFlag = 1
                       '         Debug.Print RawDataErr, WMARawData(j, i), WMAPattern(k, j, i)
                            End If
                              
                              RawDataErrSum = RawDataErrSum + RawDataErr
                             
                             
                            ' direction
                            If (WMAPattern(k, j, i) - WMAPattern(k, j, i - 1)) >= 0 Then
                            
                               PatternDirection = PatternDirection + 1
                            Else
                               PatternDirection = PatternDirection - 1
                            End If
                            
                           
                            
                            If (WMARawData(j, i) - WMARawData(j, i - 1)) >= 0 Then
                              
                               RawDataDirection = RawDataDirection + 1
                            Else
                               RawDataDirection = RawDataDirection - 1
                            End If
                           
                       End If
                        
                  
                Next i
                
      
              '  Tester.Print "WMA Pattern Dir="; PatternDirection
               ' Tester.Print "WMA RawData Dir="; RawDataDirection
                
               Tester.Print "WMA RawData Err="; RawDataErrSum
                
               
                
                If j = 0 Then
                        WMALErr = RawDataErrSum
                        DirLErr = Abs(PatternDirection - RawDataDirection)
                Else
                       WMARErr = RawDataErrSum
                       DirRErr = Abs(PatternDirection - RawDataDirection)
                End If
                
                   
                If (WMALErr < 10) And (WMARErr < 10) And j = 1 Then
              '  If DirLErr <= 2 And DirRErr <= 2 And flag = 0 Then
                        WMAMatchIndex = k ' get correct pattern
                      '  Tester.Cls
                        Tester.Print "WMA mode="; k
                        
                        Exit Sub
               ' End If
                End If
                    
           Next j
           
       
    Next k
    ' compare fail
   WMACompareResult = 2
End If
 
Exit Sub
Err1:
MsgBox "Data compare Err"





End Sub

Public Sub CompareDataAU3150JLF27MP3(AudioType As Byte)
On Error GoTo Err1
 Dim ErrCount As Integer
 Dim OldCount As Integer
 Dim CoutinueCount As Integer
 
MP3CompareResult = 1
WMACompareResult = 1

Dim k As Integer
Dim j As Integer
Dim i As Integer
Dim PatternErrSum As Double
Dim PatternErr As Single
Dim PatternDirection As Integer
Dim RawDataErrSum As Double
Dim RawDataDirection As Integer
Dim RawDataErr As Single
Dim MP3LErr As Single
Dim MP3RErr As Single
Dim DirLErr As Integer
Dim DirRErr As Integer
 
Dim WMALErr As Single
Dim WMARErr As Single
Dim FailFlag As Byte
If AudioType = 0 Then
    
    For k = 0 To MP3PatternIndex - 1
    
              FailFlag = 0
           
           For j = 0 To 1     ' channel index
                 
                  PatternDirection = 0
                   RawDataDirection = 0
                
                If Abs(MP3Pattern(k, j, 0)) > 5000 Then
                  
                 RawDataErr = ((CSng(MP3RawData(j, 0)) - CSng(MP3Pattern(k, j, 0)))) / CSng(MP3Pattern(k, j, 0))
                 RawDataErr = Abs(RawDataErr)
                 RawDataErrSum = RawDataErr
                  If RawDataErr > 0.5 Then
                             FailFlag = 1
                  End If
                Else
                 
                   
                 RawDataErrSum = 0
                End If
                
                
                For i = 1 To 99
                          Debug.Print MP3RawData(j, i)
                       
                        If Abs(MP3Pattern(k, j, i)) > 5000 Then
                          
                     
                             RawDataErr = ((CSng(MP3RawData(j, i)) - CSng(MP3Pattern(k, j, i))) / CSng(MP3Pattern(k, j, i)))
                             
                             RawDataErr = Abs(RawDataErr)
                           
                             
                            If RawDataErr > 0.5 Then
                              FailFlag = 1
                      '          Debug.Print RawDataErr, MP3RawData(j, i), MP3Pattern(k, j, i)
                            End If
                              
                              RawDataErrSum = RawDataErrSum + RawDataErr
                             
                             
                            ' direction
                            If (MP3Pattern(k, j, i) - MP3Pattern(k, j, i - 1)) >= 0 Then
                            
                               PatternDirection = PatternDirection + 1
                            Else
                               PatternDirection = PatternDirection - 1
                            End If
                            
                           
                            
                            If (MP3RawData(j, i) - MP3RawData(j, i - 1)) >= 0 Then
                              
                               RawDataDirection = RawDataDirection + 1
                            Else
                               RawDataDirection = RawDataDirection - 1
                            End If
                           
                       End If
                        
                  
                Next i
                
      
              '  Tester.Print "MP3 Pattern Dir="; PatternDirection
               ' Tester.Print "MP3 RawData Dir="; RawDataDirection
                
                Tester.Print "MP3 RawData Err="; RawDataErrSum
                
               
                
                If j = 0 Then
                        MP3LErr = RawDataErrSum
                        DirLErr = Abs(PatternDirection - RawDataDirection)
                Else
                       MP3RErr = RawDataErrSum
                       DirRErr = Abs(PatternDirection - RawDataDirection)
                End If
                
                   
                If (MP3LErr < 20) And (MP3RErr < 20) And j = 1 Then
             '   If DirLErr <= 2 And DirRErr <= 2 And flag = 0 Then
                        MP3MatchIndex = k ' get correct pattern
                      '  Tester.Cls
                        Tester.Print "MP3 mode="; k
                        
                        Exit Sub
              '  End If
                End If
                    
           Next j
           
       
    Next k
    ' compare fail
   MP3CompareResult = 2
End If
 
 
'=====================================================================
'
'   WMA process
'
'=====================================================================




If AudioType = 1 Then
    
    For k = 0 To WMAPatternIndex - 1
    
              FailFlag = 0
           
           For j = 0 To 1     ' channel index
                 
                  PatternDirection = 0
                   RawDataDirection = 0
                
                If Abs(WMAPattern(k, j, 0)) > 5000 Then
                  
                 RawDataErr = ((WMARawData(j, 0) - WMAPattern(k, j, 0))) / WMAPattern(k, j, 0)
                 RawDataErr = Abs(RawDataErr)
                 RawDataErrSum = RawDataErr
                  If RawDataErr > 0.5 Then
                             FailFlag = 1
                  End If
                Else
                 
                   
                 RawDataErrSum = 0
                End If
                
                
                For i = 1 To 99
          
                       
                        If Abs(WMAPattern(k, j, i)) > 5000 Then
                     
                             RawDataErr = ((WMARawData(j, i) - WMAPattern(k, j, i)) / WMAPattern(k, j, i))
                             
                             RawDataErr = Abs(RawDataErr)
                           
                             
                            If RawDataErr > 0.5 Then
                              FailFlag = 1
                       '         Debug.Print RawDataErr, WMARawData(j, i), WMAPattern(k, j, i)
                            End If
                              
                              RawDataErrSum = RawDataErrSum + RawDataErr
                             
                             
                            ' direction
                            If (WMAPattern(k, j, i) - WMAPattern(k, j, i - 1)) >= 0 Then
                            
                               PatternDirection = PatternDirection + 1
                            Else
                               PatternDirection = PatternDirection - 1
                            End If
                            
                           
                            
                            If (WMARawData(j, i) - WMARawData(j, i - 1)) >= 0 Then
                              
                               RawDataDirection = RawDataDirection + 1
                            Else
                               RawDataDirection = RawDataDirection - 1
                            End If
                           
                       End If
                        
                  
                Next i
                
      
              '  Tester.Print "WMA Pattern Dir="; PatternDirection
               ' Tester.Print "WMA RawData Dir="; RawDataDirection
                
                Tester.Print "WMA RawData Err="; RawDataErrSum
                
               
                
                If j = 0 Then
                        WMALErr = RawDataErrSum
                        DirLErr = Abs(PatternDirection - RawDataDirection)
                Else
                       WMARErr = RawDataErrSum
                       DirRErr = Abs(PatternDirection - RawDataDirection)
                End If
                
                   
                If (WMALErr < 10) And (WMARErr < 10) And j = 1 Then
              '  If DirLErr <= 2 And DirRErr <= 2 And flag = 0 Then
                        WMAMatchIndex = k ' get correct pattern
                      '  Tester.Cls
                        Tester.Print "WMA mode="; k
                        
                        Exit Sub
               ' End If
                End If
                    
           Next j
           
       
    Next k
    ' compare fail
   WMACompareResult = 2
End If
 
Exit Sub
Err1:
MsgBox "Data compare Err"





End Sub

Public Sub CompareData1(AudioType As Byte)
On Error Resume Next
 Dim ErrCount As Integer
 Dim OldCount As Integer
 Dim CoutinueCount As Integer
 
MP3CompareResult = 1
WMACompareResult = 1

Dim k As Integer
Dim j As Integer
Dim i As Integer
Dim PatternErrSum As Double
Dim PatternErr As Single
Dim PatternDirection As Integer
Dim RawDataErrSum As Double
Dim RawDataDirection As Integer
Dim RawDataErr As Single
Dim MP3LErr As Single
Dim MP3RErr As Single
Dim DirLErr As Integer
Dim DirRErr As Integer
 
Dim WMALErr As Single
Dim WMARErr As Single
Dim FailFlag As Byte
If AudioType = 0 Then
    
    For k = 0 To MP3PatternIndex - 1
    
              FailFlag = 0
           
           For j = 0 To 1     ' channel index
                 
                  PatternDirection = 0
                   RawDataDirection = 0
                
                If Abs(MP3Pattern(k, j, 0)) > 5000 Then
                  
                 RawDataErr = ((MP3RawData(j, 0) - MP3Pattern(k, j, 0))) / MP3Pattern(k, j, 0)
                 RawDataErr = Abs(RawDataErr)
                 RawDataErrSum = RawDataErr
                  If RawDataErr > 0.5 Then
                             FailFlag = 1
                  End If
                Else
                 
                   
                 RawDataErrSum = 0
                End If
                
                
                For i = 1 To 99
          
                       
                        If Abs(MP3Pattern(k, j, i)) > 5000 Then
                     
                             RawDataErr = ((MP3RawData(j, i) - MP3Pattern(k, j, i)) / MP3Pattern(k, j, i))
                             
                             RawDataErr = Abs(RawDataErr)
                           
                             
                            If RawDataErr > 0.5 Then
                              FailFlag = 1
                      '          Debug.Print RawDataErr, MP3RawData(j, i), MP3Pattern(k, j, i)
                            End If
                              
                              RawDataErrSum = RawDataErrSum + RawDataErr
                             
                             
                            ' direction
                            If (MP3Pattern(k, j, i) - MP3Pattern(k, j, i - 1)) >= 0 Then
                            
                               PatternDirection = PatternDirection + 1
                            Else
                               PatternDirection = PatternDirection - 1
                            End If
                            
                           
                            
                            If (MP3RawData(j, i) - MP3RawData(j, i - 1)) >= 0 Then
                              
                               RawDataDirection = RawDataDirection + 1
                            Else
                               RawDataDirection = RawDataDirection - 1
                            End If
                           
                       End If
                        
                  
                Next i
                
      
              '  Tester.Print "MP3 Pattern Dir="; PatternDirection
               ' Tester.Print "MP3 RawData Dir="; RawDataDirection
                
                Tester.Print "MP3 RawData Err="; RawDataErrSum
                
               
                
                If j = 0 Then
                        MP3LErr = RawDataErrSum
                        DirLErr = Abs(PatternDirection - RawDataDirection)
                Else
                       MP3RErr = RawDataErrSum
                       DirRErr = Abs(PatternDirection - RawDataDirection)
                End If
                
                   
                If (MP3LErr < 10) And (MP3RErr < 10) And j = 1 Then
             '   If DirLErr <= 2 And DirRErr <= 2 And flag = 0 Then
                        MP3MatchIndex = k ' get correct pattern
                      '  Tester.Cls
                        Tester.Print "MP3 mode="; k
                        
                        Exit Sub
              '  End If
                End If
                    
           Next j
           
       
    Next k
    ' compare fail
   MP3CompareResult = 2
End If
 
 
'=====================================================================
'
'   WMA process
'
'=====================================================================




If AudioType = 1 Then
    
    For k = 0 To WMAPatternIndex - 1
    
              FailFlag = 0
           
           For j = 0 To 1     ' channel index
                 
                  PatternDirection = 0
                   RawDataDirection = 0
                
                If Abs(WMAPattern(k, j, 0)) > 5000 Then
                  
                 RawDataErr = ((WMARawData(j, 0) - WMAPattern(k, j, 0))) / WMAPattern(k, j, 0)
                 RawDataErr = Abs(RawDataErr)
                 RawDataErrSum = RawDataErr
                  If RawDataErr > 0.5 Then
                             FailFlag = 1
                  End If
                Else
                 
                   
                 RawDataErrSum = 0
                End If
                
                
                For i = 1 To 99
          
                       
                        If Abs(WMAPattern(k, j, i)) > 5000 Then
                     
                             RawDataErr = ((WMARawData(j, i) - WMAPattern(k, j, i)) / WMAPattern(k, j, i))
                             
                             RawDataErr = Abs(RawDataErr)
                           
                             
                            If RawDataErr > 0.5 Then
                              FailFlag = 1
                       '         Debug.Print RawDataErr, WMARawData(j, i), WMAPattern(k, j, i)
                            End If
                              
                              RawDataErrSum = RawDataErrSum + RawDataErr
                             
                             
                            ' direction
                            If (WMAPattern(k, j, i) - WMAPattern(k, j, i - 1)) >= 0 Then
                            
                               PatternDirection = PatternDirection + 1
                            Else
                               PatternDirection = PatternDirection - 1
                            End If
                            
                           
                            
                            If (WMARawData(j, i) - WMARawData(j, i - 1)) >= 0 Then
                              
                               RawDataDirection = RawDataDirection + 1
                            Else
                               RawDataDirection = RawDataDirection - 1
                            End If
                           
                       End If
                        
                  
                Next i
                
      
              '  Tester.Print "WMA Pattern Dir="; PatternDirection
               ' Tester.Print "WMA RawData Dir="; RawDataDirection
                
                Tester.Print "WMA RawData Err="; RawDataErrSum
                
               
                
                If j = 0 Then
                        WMALErr = RawDataErrSum
                        DirLErr = Abs(PatternDirection - RawDataDirection)
                Else
                       WMARErr = RawDataErrSum
                       DirRErr = Abs(PatternDirection - RawDataDirection)
                End If
                
                   
                If (WMALErr < 10) And (WMARErr < 10) And j = 1 Then
              '  If DirLErr <= 2 And DirRErr <= 2 And flag = 0 Then
                        WMAMatchIndex = k ' get correct pattern
                      '  Tester.Cls
                        Tester.Print "WMA mode="; k
                        
                        Exit Sub
               ' End If
                End If
                    
           Next j
           
       
    Next k
    ' compare fail
   WMACompareResult = 2
End If
 

Err1:
MsgBox "Data compare Err"





End Sub


Public Sub CompareData2(AudioType As Byte)
 Dim ErrCount As Integer
 Dim OldCount As Integer
 Dim CoutinueCount As Integer
 
MP3CompareResult = 1
WMACompareResult = 1

Dim k As Integer
Dim j As Integer
Dim i As Integer
 
 
If AudioType = 0 Then
    
    For k = 0 To MP3PatternIndex - 1
    
          
           
           For j = 0 To 1     ' channel index
           
                ErrCount = 0
                CoutinueCount = 0
           
                For i = 0 To 99
                
                    If CSng(MP3RawData(j, i)) = 0 Then
                       MP3RawData(j, i) = 1
                    End If
                   
                   If CSng(MP3Pattern(k, j, i)) / CSng(MP3RawData(j, i)) < 0.7 Or CSng(MP3Pattern(k, j, i)) / CSng(MP3RawData(j, i)) > 1.3 Then
                     
                         ErrCount = ErrCount + 1
                     
                         If i = OldCount + 1 Then   '  continue fail condition
                             CoutinueCount = CoutinueCount + 1
                         End If
                    End If
                 
                  OldCount = i
                Next i
                
                Tester.Print "MP3 ERR="; ErrCount
                If ErrCount < 20 Then
                        MP3MatchIndex = k ' get correct pattern
                          Tester.Print "MP3 mode="; k
                        Exit Sub
                End If
           
           Next j
       
    Next k
    ' compare fail
   MP3CompareResult = 2
End If
 
 
'=====================================================================
'
'   WMA process
'
'=====================================================================




If AudioType = 1 Then
    
    For k = 0 To WMAPatternIndex - 1
    
           
           
           For j = 0 To 1     ' MP3 data
           
                ErrCount = 0
                CoutinueCount = 0
           
                For i = 0 To 99
                
                    If CSng(WMARawData(j, i)) = 0 Then
                       WMARawData(j, i) = 1
                    End If
                   
                   If CSng(WMAPattern(k, j, i)) / CSng(WMARawData(j, i)) < 0.7 Or CSng(WMAPattern(k, j, i)) / CSng(WMARawData(j, i)) > 1.3 Then
                     
                         ErrCount = ErrCount + 1
                     
                         If i = OldCount + 1 Then   '  continue fail condition
                             CoutinueCount = CoutinueCount + 1
                         End If
                    End If
                 
                  OldCount = i
                Next i
                Tester.Print "WMA ERR="; ErrCount
                
                If ErrCount < 20 Then
                        WMAMatchIndex = k
                         Tester.Print "WMA mode="; k
                        Exit Sub
                End If
           
           Next j
       
    Next k
    WMACompareResult = 2
End If







End Sub

Public Sub TwoChannelTest()

               If ChipName = "AU3150KLF24" Then
                Call AU3150K_1_SlotTest_Noise
                End If
             
                If ChipName = "AU3152CLF24" Then
                Call AU3152CLF24_1_SlotTest_Noise
                End If
                
                If ChipName = "AU3150JLF27" Then
                Call AU3150JLF27_1_SlotTest_Noise
                End If
           
           
                 If ChipName = "AU3150KLF27" Then
                Call AU3150KLF27_1_SlotTest_Noise
                End If
                  If ChipName = "AU3152ALF24" Then
                Call AU3152ALF24_1_SlotTest_Noise
                End If
End Sub


Public Sub MP3Tester()


    'Call AU3130_1_SlotTest_WMA
                If ChipName = "AU3150KLF24" Then
                    Call AU3150K_1_SlotTest_Noise
                End If
             
                If ChipName = "AU3152CLF24" Then
                    Call AU3152CLF24_1_SlotTest_Noise
                End If
                
                
                 If ChipName = "AU3152CLF27" Then
                    Call AU3152CLF27_1_SlotTest_Noise
                End If
                
                  
                If ChipName = "AU3152ALF24" Then
                    Call AU3152ALF24_1_SlotTest_Noise
                End If
                
                
                If ChipName = "AU3150ALF24" Then
                    Call AU3150ALF24_2_SlotTest_Noise
                End If

                
                If ChipName = "AU3150JLF27" Then
                    Call AU3150JLF27_1_SlotTest_Noise
                End If
                
        
                
                 
                If ChipName = "AU3150KLF27" Then
                    Call AU3150KLF27_1_SlotTest_Noise
                End If
           
           
                  If ChipName = "AU3150NLF27" Then
                    Call AU3150NLF27_1_SlotTest_Noise
                End If
                
                   If ChipName = "AU3150PLF27" Then
                    Call AU3150PLF27_1_SlotTest_Noise
                End If
                
                   If ChipName = "AU3150QLF27" Then
                    Call AU3150QLF27_1_SlotTest_Noise
                End If
                
                      If ChipName = "AU3152ALF27" Then
                    Call AU3152ALF27_1_SlotTest_Noise
                End If
                
                
                If ChipName = "AU3150ALF27" Then
                    Call AU3150ALF27_2_SlotTest_Noise
                End If
                
                If ChipName = "AU3150LLF27" Then
                    Call AU3150LLF27_1_SlotTest_Noise
                End If
               
               
                If ChipName = "AU3150CLF27" Then
                      Call AU3150CLF27_2_SlotTest_Noise
                End If
                
                
                 If ChipName = "AU3150CLF2C" Then
                      Call AU3150CLF2C_2_SlotTest_Noise
                End If
                '=================================
                
                If ChipName = "AU3150ALF28" Then
                    Call AU3150ALF28_2_SlotTest_Noise
                End If
                
              '  If ChipName = "AU3150CLF28" Then
              '      Call AU3150CLF28_2_SlotTest_Noise
              '  End If
                
                If ChipName = "AU3150LLF28" Then
                    Call AU3150LLF28_1_SlotTest_Noise
                End If
                
                   If ChipName = "AU3150LLF2C" Then
                    Call AU3150LLF2C_1_SlotTest_Noise
                End If
                
                
                If ChipName = "AU3150LLF2D" Then
                    Call AU3150LLF2D_1_SlotTest_Noise
                End If
                
                If ChipName = "AU3150JLF28" Then
                    Call AU3150JLF28_1_SlotTest_Noise
                End If
                
                 If ChipName = "AU3150KLF28" Then
                    Call AU3150KLF28_1_SlotTest_Noise
                End If
                
             
                
                                
                If ChipName = "AU3150NLF28" Then
                    Call AU3150NLF28_1_SlotTest_Noise
                End If
                
                 If ChipName = "AU3150PLF28" Then
                    Call AU3150PLF28_1_SlotTest_Noise
                End If
                
                 If ChipName = "AU3150QLF28" Then
                    Call AU3150QLF28_1_SlotTest_Noise
                End If
                
             
                If ChipName = "AU3152ALF28" Then
                    Call AU3152ALF28_1_SlotTest_Noise
                End If
               
                
                 If ChipName = "AU3152CLF28" Then
                    Call AU3152CLF28_1_SlotTest_Noise
                End If
                
                '====================
                
                 If ChipName = "AU3152HLF2C" Then
                    Call AU3152HLF2C_1_SlotTest_Noise
                End If
                
                
                   If ChipName = "AU3152ALF2C" Then
                    Call AU3152ALF2C_1_SlotTest_Noise
                End If
                
                
                 If ChipName = "AU3152ALF2D" Then
                    Call AU3152ALF2D_1_SlotTest_Noise
                End If
                
               
                
                
                   If ChipName = "AU3152HLF2D" Then
                    Call AU3152HLF2D_1_SlotTest_Noise
                End If
                
                    If ChipName = "AU3152HLF2E" Then
                    Call AU3152HLF2E_1_SlotTest_Noise
                End If
                
End Sub




Public Sub ADCBuffer2AudioBuffer(AudioType As Byte)
 
Dim i As Integer
 
If AudioType = 0 Then   'MP3 type

        For i = 0 To 5999
        
            MP3RawData(0, i) = ADCBuffer(i * 2)
            MP3RawData(1, i) = ADCBuffer(i * 2 + 1)
         
        Next i
        
 End If
 
 
 
 If AudioType = 1 Then   'WMA type

        For i = 0 To 5999
        
            WMARawData(0, i) = ADCBuffer(i * 2)
            WMARawData(1, i) = ADCBuffer(i * 2 + 1)
         
        Next i
        
 End If







End Sub
Public Function DACGetData2ChannelRange25() As Integer
'On Error Resume Next
  Dim result As Long
  Dim Status As Byte
  Dim i As Long, k As Long
  Dim ReTrgCnt As Integer, MCnt As Integer, PostCnt As Integer
  Dim TimerSrc As Integer
  Dim cobTrigModeListIndex As Integer
  Dim cobTrigSrcListIndex As Integer
  Dim cobTrigPolListIndex As Integer
  Dim cobClkSrcListIndex As Integer
  Dim ErrCount As Integer
  Dim OldCount As Integer
  Dim CoutinueCount As Integer
  Dim cobTrigLevelText As String
  Dim cobTrigL_LevelText As String
  Dim gnClkDiv As Integer
  Dim Id As Integer
  Dim startPos As Long
  Dim OldTimer
  Dim ADCPostCnt As Integer
  Dim ADCdma_size As Long
  
   ' dma2205.Hide
  ADCPostCnt = 1 ' post trigger 500
  ADCdma_size = 100 ' scan count 2000-- 80000
  ch_cnt = 2
  ADCch_num(0) = 0  ' channel 0
  
  ADCch_range(0) = 3  ' BiPolar 2.5V, 1 base
  ADCch_ref(0) = 1    ' Differentail
  
  
  ADCch_num(1) = 1  ' channel 0
  
  
   ADCch_range(1) = 3  ' BiPolar 2.5V, 1 base
  ADCch_ref(0) = 1    ' Differentail
  
  '=================================
  '  DMA setting
  '=================================
  
        cobTrigModeListIndex = 0
        cobTrigSrcListIndex = 1
        cobTrigPolListIndex = 1
        cobClkSrcListIndex = 0
        cobTrigPolListIndex = 1
        cobTrigLevelText = "200"
        cobTrigL_LevelText = "128"
        gnClkDiv = 5000
   
       
    Call MsecDelay(0.5)   'for AU3150 unstable chip
    

   If Card2214Exist = 0 Then
        Card2214 = D2K_Register_Card(DAQ_2214, 0)
        
          For i = 0 To 1 Step 1
                result = D2K_AI_CH_Config(Card2214, ADCch_num(i), ADCch_range(i) Or (ADCch_ref(i) * 256))
          Next
          
   
          
         
    End If
   
   
  For i = 0 To 99
  ADCBuffer(i) = 0
  Next i
  
   
        result = D2K_AI_Config(Card2214, 0, (CLng(cobTrigModeListIndex) * 8) Or CLng(cobTrigSrcListIndex) Or (CLng(cobTrigPolListIndex) * 4096), ADCPostCnt, 0, 1, 1)
          If result <> 0 Then
            MsgBox "2214 configure fail 1"
            Exit Function
          End If
   
   
  
   result = D2K_AI_ContBufferSetup(Card2214, ADCBuffer(0), ADCdma_size * ch_cnt, Id)
'   Tester.Print "result2 ="; result; "Id="; Id
   
    result = D2K_AIO_Config(Card2214, cobClkSrcListIndex, (CLng(cobTrigPolListIndex) * 256) Or CH0ATRIG, CLng(cobTrigLevelText), CLng(cobTrigL_LevelText))
            If result <> 0 Then
            MsgBox "2214 configure fail 2"
            Exit Function
            End If
  'Tester.Print "result4="; result
  result = D2K_AI_ContReadMultiChannels(Card2214, ch_cnt, ADCch_num(0), Id, ADCdma_size, gnClkDiv * ch_cnt, gnClkDiv, ASYNCH_OP)
     
 '  Tester.Print "read AD begin ="; result
   Card2214Exist = 1
  
  
  
  
  Status = 0
  bStop = 0
  OldTimer = Timer
  While Status = 0 And bStop = 0 And (Timer - OldTimer < 10)
     DoEvents
    result = D2K_AI_AsyncCheck(Card2214, Status, count1)
    
  Wend
  result = D2K_AI_AsyncClear(Card2214, startPos, count1)
  Tester.Print "result5 ="; result;
   Tester.Print "read AD end "
  
 End Function
  Public Function DACGetData2ChannelUpTriggle() As Integer
'On Error Resume Next
  Dim result As Long
  Dim Status As Byte
  Dim i As Long, k As Long
  Dim ReTrgCnt As Integer, MCnt As Integer, PostCnt As Integer
  Dim TimerSrc As Integer
  Dim cobTrigModeListIndex As Integer
  Dim cobTrigSrcListIndex As Integer
  Dim cobTrigPolListIndex As Integer
  Dim cobClkSrcListIndex As Integer
  Dim ErrCount As Integer
  Dim OldCount As Integer
  Dim CoutinueCount As Integer
  Dim cobTrigLevelText As String
  Dim cobTrigL_LevelText As String
  Dim gnClkDiv As Integer
  Dim Id As Integer
  Dim startPos As Long
  Dim OldTimer
  Dim ADCPostCnt As Integer
  Dim ADCdma_size As Long
  
   ' dma2205.Hide
  ADCPostCnt = 5000 'post trigger 500
  ADCdma_size = 6000 ' scan count 2000-- 80000
  ch_cnt = 2
  ADCch_num(0) = 0  ' channel 0
  ADCch_range(0) = 4  ' BiPolar 1.25V, 1 base
 
  ADCch_ref(0) = 1    ' Differentail
  
  
  ADCch_num(1) = 1  ' channel 0
  ADCch_range(1) = 4  ' BiPolar 1.25V, 1 base
  
  ADCch_ref(0) = 1    ' Differentail
  
  '=================================
  '  DMA setting
  '=================================
  
        cobTrigModeListIndex = 0
        cobTrigSrcListIndex = 1
        cobTrigPolListIndex = 1
        cobClkSrcListIndex = 0
        cobTrigPolListIndex = 1   ' because the layout has error at right-left way
        cobTrigLevelText = "200"
        cobTrigL_LevelText = "20"  ' because the layout has error
        gnClkDiv = 5000
   
       
    Call MsecDelay(0.5)   'for AU3150 unstable chip
  
   If Card2214Exist = 0 Then
        Card2214 = D2K_Register_Card(DAQ_2214, 0)
        
          For i = 0 To 1 Step 1
                result = D2K_AI_CH_Config(Card2214, ADCch_num(i), ADCch_range(i) Or (ADCch_ref(i) * 256))
          Next
          
   
          
         
    End If
   
   
  For i = 0 To 11999
  ADCBuffer(i) = 0
  Next i
  
   
        result = D2K_AI_Config(Card2214, 0, (CLng(cobTrigModeListIndex) * 8) Or CLng(cobTrigSrcListIndex) Or (CLng(cobTrigPolListIndex) * 4096), ADCPostCnt, 0, 1, 1)
          If result <> 0 Then
            MsgBox "2214 configure  fail 1"
            Exit Function
          End If
   
    
 
    result = D2K_AI_ContBufferSetup(Card2214, ADCBuffer(0), ADCdma_size * ch_cnt, Id)
'   Tester.Print "result2 ="; result; "Id="; Id
  If result <> 0 Then
            MsgBox "2214 configure  fail 2"
            Exit Function
          End If
   

   
    result = D2K_AIO_Config(Card2214, cobClkSrcListIndex, (CLng(cobTrigPolListIndex) * 256) Or CH0ATRIG, CLng(cobTrigLevelText), CLng(cobTrigL_LevelText))
            If result <> 0 Then
            MsgBox "2214 configure fail 3"
            Exit Function
            End If
  'Tester.Print "result4="; result
  
  result = D2K_AI_ContReadMultiChannels(Card2214, ch_cnt, ADCch_num(0), Id, ADCdma_size, gnClkDiv * ch_cnt, gnClkDiv, ASYNCH_OP)
       If result <> 0 Then
            MsgBox "2214 configure fail 4"
            Exit Function
          End If
   
 '  Tester.Print "read AD begin ="; result
   Card2214Exist = 1
  
  
  
  
  Status = 0
  bStop = 0
  OldTimer = Timer
  While Status = 0 And bStop = 0 And (Timer - OldTimer < 10)
     DoEvents
    result = D2K_AI_AsyncCheck(Card2214, Status, count1)
    
  Wend
  result = D2K_AI_AsyncClear(Card2214, startPos, count1)
  Tester.Print "result5 ="; result;
   Tester.Print "read AD end "
  
 End Function
 
 Public Function DACGetData2ChannelUpTriggleAU3152HL() As Integer
'On Error Resume Next
  Dim result As Long
  Dim Status As Byte
  Dim i As Long, k As Long
  Dim ReTrgCnt As Integer, MCnt As Integer, PostCnt As Integer
  Dim TimerSrc As Integer
  Dim cobTrigModeListIndex As Integer
  Dim cobTrigSrcListIndex As Integer
  Dim cobTrigPolListIndex As Integer
  Dim cobClkSrcListIndex As Integer
  Dim ErrCount As Integer
  Dim OldCount As Integer
  Dim CoutinueCount As Integer
  Dim cobTrigLevelText As String
  Dim cobTrigL_LevelText As String
  Dim gnClkDiv As Integer
  Dim Id As Integer
  Dim startPos As Long
  Dim OldTimer
  Dim ADCPostCnt As Integer
  Dim ADCdma_size As Long
  
   ' dma2205.Hide
  ADCPostCnt = 5000 'post trigger 500
  ADCdma_size = 6000 ' scan count 2000-- 80000
  ch_cnt = 2
  ADCch_num(0) = 0  ' channel 0
  ADCch_range(0) = 4  ' BiPolar 1.25V, 1 base
 
  ADCch_ref(0) = 1    ' Differentail
  
  
  ADCch_num(1) = 1  ' channel 0
  ADCch_range(1) = 4  ' BiPolar 1.25V, 1 base
  
  ADCch_ref(0) = 1    ' Differentail
  
  '=================================
  '  DMA setting
  '=================================
  
      '=================================
  '  DMA setting
  '=================================
  
        cobTrigModeListIndex = 0
        cobTrigSrcListIndex = 1
        cobTrigPolListIndex = 1
        cobClkSrcListIndex = 0
        cobTrigPolListIndex = 1   ' because the layout has error at right-left way
        cobTrigLevelText = "150"
        cobTrigL_LevelText = "20"  ' because the layout has error
        gnClkDiv = 5000
   
       
    Call MsecDelay(0.5)   'for AU3150 unstable chip
  
   If Card2214Exist = 0 Then
        Card2214 = D2K_Register_Card(DAQ_2214, 0)
        
          For i = 0 To 1 Step 1
                result = D2K_AI_CH_Config(Card2214, ADCch_num(i), ADCch_range(i) Or (ADCch_ref(i) * 256))
          Next
          
   
          
         
    End If
   
   
  For i = 0 To 11999
  ADCBuffer(i) = 0
  Next i
  
   
        result = D2K_AI_Config(Card2214, 0, (CLng(cobTrigModeListIndex) * 8) Or CLng(cobTrigSrcListIndex) Or (CLng(cobTrigPolListIndex) * 4096), ADCPostCnt, 0, 1, 1)
          If result <> 0 Then
            MsgBox "2214 configure  fail 1"
            Exit Function
          End If
   
    
 
    result = D2K_AI_ContBufferSetup(Card2214, ADCBuffer(0), ADCdma_size * ch_cnt, Id)
'   Tester.Print "result2 ="; result; "Id="; Id
  If result <> 0 Then
            MsgBox "2214 configure  fail 2"
            Exit Function
          End If
   

   
    result = D2K_AIO_Config(Card2214, cobClkSrcListIndex, (CLng(cobTrigPolListIndex) * 256) Or CH0ATRIG, CLng(cobTrigLevelText), CLng(cobTrigL_LevelText))
            If result <> 0 Then
            MsgBox "2214 configure fail 3"
            Exit Function
            End If
  'Tester.Print "result4="; result
  
  result = D2K_AI_ContReadMultiChannels(Card2214, ch_cnt, ADCch_num(0), Id, ADCdma_size, gnClkDiv * ch_cnt, gnClkDiv, ASYNCH_OP)
       If result <> 0 Then
            MsgBox "2214 configure fail 4"
            Exit Function
          End If
   
 '  Tester.Print "read AD begin ="; result
   Card2214Exist = 1
  
  
  
  
  Status = 0
  bStop = 0
  OldTimer = Timer
  While Status = 0 And bStop = 0 And (Timer - OldTimer < 10)
     DoEvents
    result = D2K_AI_AsyncCheck(Card2214, Status, count1)
    
  Wend
  result = D2K_AI_AsyncClear(Card2214, startPos, count1)
  Tester.Print "result5 ="; result;
   Tester.Print "read AD end "
  
 End Function
  Public Function DACGetData2ChannelLowTriggle() As Integer
'On Error Resume Next
  Dim result As Long
  Dim Status As Byte
  Dim i As Long, k As Long
  Dim ReTrgCnt As Integer, MCnt As Integer, PostCnt As Integer
  Dim TimerSrc As Integer
  Dim cobTrigModeListIndex As Integer
  Dim cobTrigSrcListIndex As Integer
  Dim cobTrigPolListIndex As Integer
  Dim cobClkSrcListIndex As Integer
  Dim ErrCount As Integer
  Dim OldCount As Integer
  Dim CoutinueCount As Integer
  Dim cobTrigLevelText As String
  Dim cobTrigL_LevelText As String
  Dim gnClkDiv As Integer
  Dim Id As Integer
  Dim startPos As Long
  Dim OldTimer
  Dim ADCPostCnt As Integer
  Dim ADCdma_size As Long
  
   ' dma2205.Hide
  ADCPostCnt = 500 ' post trigger 500
  ADCdma_size = 100 ' scan count 2000-- 80000
  ch_cnt = 2
  ADCch_num(0) = 0  ' channel 0
  ADCch_range(0) = 4  ' BiPolar 1.25V, 1 base
 
  ADCch_ref(0) = 1    ' Differentail
  
  
  ADCch_num(1) = 1  ' channel 0
  ADCch_range(1) = 4  ' BiPolar 1.25V, 1 base
  
  ADCch_ref(0) = 1    ' Differentail
  
  '=================================
  '  DMA setting
  '=================================
  
        cobTrigModeListIndex = 0
        cobTrigSrcListIndex = 1
        cobTrigPolListIndex = 1
        cobClkSrcListIndex = 0
        cobTrigPolListIndex = 0
        cobTrigLevelText = "200"
        cobTrigL_LevelText = "128"
        gnClkDiv = 5000
   
       
    Call MsecDelay(0.5)   'for AU3150 unstable chip
    

   If Card2214Exist = 0 Then
        Card2214 = D2K_Register_Card(DAQ_2214, 0)
        
          For i = 0 To 1 Step 1
                result = D2K_AI_CH_Config(Card2214, ADCch_num(i), ADCch_range(i) Or (ADCch_ref(i) * 256))
          Next
          
   
          
         
    End If
   
   
  For i = 0 To 99
  ADCBuffer(i) = 0
  Next i
  
   
        result = D2K_AI_Config(Card2214, 0, (CLng(cobTrigModeListIndex) * 8) Or CLng(cobTrigSrcListIndex) Or (CLng(cobTrigPolListIndex) * 4096), ADCPostCnt, 0, 1, 1)
          If result <> 0 Then
            MsgBox "2214 configure fail 1"
            Exit Function
          End If
   
   
  
   result = D2K_AI_ContBufferSetup(Card2214, ADCBuffer(0), ADCdma_size * ch_cnt, Id)
'   Tester.Print "result2 ="; result; "Id="; Id
   
    result = D2K_AIO_Config(Card2214, cobClkSrcListIndex, (CLng(cobTrigPolListIndex) * 256) Or CH0ATRIG, CLng(cobTrigLevelText), CLng(cobTrigL_LevelText))
            If result <> 0 Then
            MsgBox "2214 configure fail 2"
            Exit Function
            End If
  'Tester.Print "result4="; result
  result = D2K_AI_ContReadMultiChannels(Card2214, ch_cnt, ADCch_num(0), Id, ADCdma_size, gnClkDiv * ch_cnt, gnClkDiv, ASYNCH_OP)
     
 '  Tester.Print "read AD begin ="; result
   Card2214Exist = 1
  
  
  
  
  Status = 0
  bStop = 0
  OldTimer = Timer
  While Status = 0 And bStop = 0 And (Timer - OldTimer < 10)
     DoEvents
    result = D2K_AI_AsyncCheck(Card2214, Status, count1)
    
  Wend
  result = D2K_AI_AsyncClear(Card2214, startPos, count1)
  Tester.Print "result5 ="; result;
   Tester.Print "read AD end "
  
 End Function
 
   Public Function DACGetData2ChannelLowTriggle64() As Integer
'On Error Resume Next
  Dim result As Long
  Dim Status As Byte
  Dim i As Long, k As Long
  Dim ReTrgCnt As Integer, MCnt As Integer, PostCnt As Integer
  Dim TimerSrc As Integer
  Dim cobTrigModeListIndex As Integer
  Dim cobTrigSrcListIndex As Integer
  Dim cobTrigPolListIndex As Integer
  Dim cobClkSrcListIndex As Integer
  Dim ErrCount As Integer
  Dim OldCount As Integer
  Dim CoutinueCount As Integer
  Dim cobTrigLevelText As String
  Dim cobTrigL_LevelText As String
  Dim gnClkDiv As Integer
  Dim Id As Integer
  Dim startPos As Long
  Dim OldTimer
  Dim ADCPostCnt As Integer
  Dim ADCdma_size As Long
  
   ' dma2205.Hide
  ADCPostCnt = 500 ' post trigger 500
  ADCdma_size = 100 ' scan count 2000-- 80000
  ch_cnt = 2
  ADCch_num(0) = 0  ' channel 0
  ADCch_range(0) = 4  ' BiPolar 1.25V, 1 base
 
  ADCch_ref(0) = 1    ' Differentail
  
  
  ADCch_num(1) = 1  ' channel 0
  ADCch_range(1) = 4  ' BiPolar 1.25V, 1 base
  
  ADCch_ref(0) = 1    ' Differentail
  
  '=================================
  '  DMA setting
  '=================================
  
        cobTrigModeListIndex = 0
        cobTrigSrcListIndex = 1
        cobTrigPolListIndex = 1
        cobClkSrcListIndex = 0
        cobTrigPolListIndex = 0
        cobTrigLevelText = "200"
        cobTrigL_LevelText = "96"
        gnClkDiv = 5000
   
       
    Call MsecDelay(0.5)   'for AU3150 unstable chip
    

   If Card2214Exist = 0 Then
        Card2214 = D2K_Register_Card(DAQ_2214, 0)
        
          For i = 0 To 1 Step 1
                result = D2K_AI_CH_Config(Card2214, ADCch_num(i), ADCch_range(i) Or (ADCch_ref(i) * 256))
          Next
          
   
          
         
    End If
   
   
  For i = 0 To 99
  ADCBuffer(i) = 0
  Next i
  
   
        result = D2K_AI_Config(Card2214, 0, (CLng(cobTrigModeListIndex) * 8) Or CLng(cobTrigSrcListIndex) Or (CLng(cobTrigPolListIndex) * 4096), ADCPostCnt, 0, 1, 1)
          If result <> 0 Then
            MsgBox "2214 configure fail 1"
            Exit Function
          End If
   
   
  
   result = D2K_AI_ContBufferSetup(Card2214, ADCBuffer(0), ADCdma_size * ch_cnt, Id)
'   Tester.Print "result2 ="; result; "Id="; Id
   
    result = D2K_AIO_Config(Card2214, cobClkSrcListIndex, (CLng(cobTrigPolListIndex) * 256) Or CH0ATRIG, CLng(cobTrigLevelText), CLng(cobTrigL_LevelText))
            If result <> 0 Then
            MsgBox "2214 configure fail 2"
            Exit Function
            End If
  'Tester.Print "result4="; result
  result = D2K_AI_ContReadMultiChannels(Card2214, ch_cnt, ADCch_num(0), Id, ADCdma_size, gnClkDiv * ch_cnt, gnClkDiv, ASYNCH_OP)
     
 '  Tester.Print "read AD begin ="; result
   Card2214Exist = 1
  
  
  
  
  Status = 0
  bStop = 0
  OldTimer = Timer
  While Status = 0 And bStop = 0 And (Timer - OldTimer < 10)
     DoEvents
    result = D2K_AI_AsyncCheck(Card2214, Status, count1)
    
  Wend
  result = D2K_AI_AsyncClear(Card2214, startPos, count1)
  Tester.Print "result5 ="; result;
   Tester.Print "read AD end "
  
 End Function
 
 Public Function DACGetData1Channel() As Integer
'On Error Resume Next
  Dim result As Long
  Dim Status As Byte
  Dim i As Long, k As Long
  Dim ReTrgCnt As Integer, MCnt As Integer, PostCnt As Integer
  Dim TimerSrc As Integer
  Dim cobTrigModeListIndex As Integer
  Dim cobTrigSrcListIndex As Integer
  Dim cobTrigPolListIndex As Integer
  Dim cobClkSrcListIndex As Integer
  Dim ErrCount As Integer
  Dim OldCount As Integer
  Dim CoutinueCount As Integer
  Dim cobTrigLevelText As String
  Dim cobTrigL_LevelText As String
  Dim gnClkDiv As Integer
  Dim Id As Integer
  Dim startPos As Long
  Dim OldTimer
  Dim ADCPostCnt As Integer
  Dim ADCdma_size As Long
  
   ' dma2205.Hide
  ADCPostCnt = 1 ' post trigger 500
  ADCdma_size = 100 ' scan count 2000-- 80000
  ch_cnt = 1    ' 1 channel
  ADCch_num(0) = 0  ' channel 0
  ADCch_range(0) = 2  ' BiPolar 5V, 1 base
 
  ADCch_ref(0) = 1    ' Differentail
  
  
  
  '=================================
  '  DMA setting
  '=================================
  
        cobTrigModeListIndex = 0 ' post trig
        cobTrigSrcListIndex = 0  ' software
        cobTrigPolListIndex = 1
        cobClkSrcListIndex = 0
        cobTrigPolListIndex = 1
        cobTrigLevelText = "200"
        cobTrigL_LevelText = "128"
        gnClkDiv = 5000
   
       
    Call MsecDelay(0.5)   'for AU3150 unstable chip
    

   If Card2214Exist = 0 Then
        Card2214 = D2K_Register_Card(DAQ_2214, 0)
        
          For i = 0 To 0 Step 1
                result = D2K_AI_CH_Config(Card2214, ADCch_num(i), ADCch_range(i) Or (ADCch_ref(i) * 256))
          Next
          
   
          
         
    End If
   
   
  For i = 0 To 99
  ADCBuffer(i) = 0
  Next i
  
   
        result = D2K_AI_Config(Card2214, 0, (CLng(cobTrigModeListIndex) * 8) Or CLng(cobTrigSrcListIndex) Or (CLng(cobTrigPolListIndex) * 4096), ADCPostCnt, 0, 1, 1)
          If result <> 0 Then
            MsgBox "2214 configure fail 1"
            Exit Function
          End If
   
   
  
   result = D2K_AI_ContBufferSetup(Card2214, ADCBuffer(0), ADCdma_size * ch_cnt, Id)
'   Tester.Print "result2 ="; result; "Id="; Id
   
    result = D2K_AIO_Config(Card2214, cobClkSrcListIndex, (CLng(cobTrigPolListIndex) * 256) Or CH0ATRIG, CLng(cobTrigLevelText), CLng(cobTrigL_LevelText))
            If result <> 0 Then
            MsgBox "2214 configure fail 2"
            Exit Function
            End If
  'Tester.Print "result4="; result
  result = D2K_AI_ContReadMultiChannels(Card2214, ch_cnt, ADCch_num(0), Id, ADCdma_size, gnClkDiv * ch_cnt, gnClkDiv, ASYNCH_OP)
     
 '  Tester.Print "read AD begin ="; result
   Card2214Exist = 1
  
  
  
  
  Status = 0
  bStop = 0
  OldTimer = Timer
  While Status = 0 And bStop = 0 And (Timer - OldTimer < 10)
     DoEvents
    result = D2K_AI_AsyncCheck(Card2214, Status, count1)
    
  Wend
  result = D2K_AI_AsyncClear(Card2214, startPos, count1)
  Tester.Print "result5 ="; result;
   Tester.Print "read AD end "
  
 End Function
  Public Function DACGetData2Channel() As Integer
'On Error Resume Next
  Dim result As Long
  Dim Status As Byte
  Dim i As Long, k As Long
  Dim ReTrgCnt As Integer, MCnt As Integer, PostCnt As Integer
  Dim TimerSrc As Integer
  Dim cobTrigModeListIndex As Integer
  Dim cobTrigSrcListIndex As Integer
  Dim cobTrigPolListIndex As Integer
  Dim cobClkSrcListIndex As Integer
  Dim ErrCount As Integer
  Dim OldCount As Integer
  Dim CoutinueCount As Integer
  Dim cobTrigLevelText As String
  Dim cobTrigL_LevelText As String
  Dim gnClkDiv As Integer
  Dim Id As Integer
  Dim startPos As Long
  Dim OldTimer
  Dim ADCPostCnt As Integer
  Dim ADCdma_size As Long
  
   ' dma2205.Hide
  ADCPostCnt = 1 ' post trigger 500
  ADCdma_size = 100 ' scan count 2000-- 80000
  ch_cnt = 2
  ADCch_num(0) = 0  ' channel 0
  ADCch_range(0) = 4  ' BiPolar 1.25V, 1 base
 
  ADCch_ref(0) = 1    ' Differentail
  
  
  ADCch_num(1) = 1  ' channel 0
  ADCch_range(1) = 4  ' BiPolar 1.25V, 1 base
  
  ADCch_ref(0) = 1    ' Differentail
  
  '=================================
  '  DMA setting
  '=================================
  
        cobTrigModeListIndex = 0
        cobTrigSrcListIndex = 1
        cobTrigPolListIndex = 1
        cobClkSrcListIndex = 0
        cobTrigPolListIndex = 1
        cobTrigLevelText = "200"
        cobTrigL_LevelText = "128"
        gnClkDiv = 5000
   
       
    Call MsecDelay(0.5)   'for AU3150 unstable chip
    

   If Card2214Exist = 0 Then
        Card2214 = D2K_Register_Card(DAQ_2214, 0)
        
          For i = 0 To 1 Step 1
                result = D2K_AI_CH_Config(Card2214, ADCch_num(i), ADCch_range(i) Or (ADCch_ref(i) * 256))
          Next
          
   
          
         
    End If
   
   
  For i = 0 To 99
  ADCBuffer(i) = 0
  Next i
  
   
        result = D2K_AI_Config(Card2214, 0, (CLng(cobTrigModeListIndex) * 8) Or CLng(cobTrigSrcListIndex) Or (CLng(cobTrigPolListIndex) * 4096), ADCPostCnt, 0, 1, 1)
          If result <> 0 Then
            MsgBox "2214 configure fail 1"
            Exit Function
          End If
   
   
  
   result = D2K_AI_ContBufferSetup(Card2214, ADCBuffer(0), ADCdma_size * ch_cnt, Id)
'   Tester.Print "result2 ="; result; "Id="; Id
   
    result = D2K_AIO_Config(Card2214, cobClkSrcListIndex, (CLng(cobTrigPolListIndex) * 256) Or CH0ATRIG, CLng(cobTrigLevelText), CLng(cobTrigL_LevelText))
            If result <> 0 Then
            MsgBox "2214 configure fail 2"
            Exit Function
            End If
  'Tester.Print "result4="; result
  result = D2K_AI_ContReadMultiChannels(Card2214, ch_cnt, ADCch_num(0), Id, ADCdma_size, gnClkDiv * ch_cnt, gnClkDiv, ASYNCH_OP)
     
 '  Tester.Print "read AD begin ="; result
   Card2214Exist = 1
  
  
  
  
  Status = 0
  bStop = 0
  OldTimer = Timer
  While Status = 0 And bStop = 0 And (Timer - OldTimer < 10)
     DoEvents
    result = D2K_AI_AsyncCheck(Card2214, Status, count1)
    
  Wend
  result = D2K_AI_AsyncClear(Card2214, startPos, count1)
  Tester.Print "result5 ="; result;
   Tester.Print "read AD end "
  
 End Function

Public Function AU3130TestMP3() As Integer
On Error Resume Next
  Dim result As Long
  Dim Status As Byte
  Dim i As Long, k As Long
  Dim ReTrgCnt As Integer, MCnt As Integer, PostCnt As Integer
  Dim TimerSrc As Integer
  Dim cobTrigModeListIndex As Integer
  Dim cobTrigSrcListIndex As Integer
  Dim cobTrigPolListIndex As Integer
  Dim cobClkSrcListIndex As Integer
  Dim ErrCount As Integer
  Dim OldCount As Integer
  Dim CoutinueCount As Integer
  Dim cobTrigLevelText As String
  Dim cobTrigL_LevelText As String
  Dim gnClkDiv As Integer
  Dim Id As Integer
  Dim startPos As Long
  Dim OldTimer
  
   ' dma2205.Hide
  PostCnt = 500 ' post trigger 500
  dma_size = 2000 ' scan count 2000-- 80000
  ch_cnt = 1
  ch_num = 0  ' channel 0
  ch_range = 4  ' BiPolar 1.25V
  ch_ref = 1    ' Differentail
  
  If Left(ChipName, 8) = "AU3130CL" Or Left(ChipName, 8) = "AU3150JL" Or ChipName = "AU3150ALF22" Then
  
        cobTrigModeListIndex = 0
        cobTrigSrcListIndex = 1
        cobTrigPolListIndex = 1
        cobClkSrcListIndex = 0
        cobTrigPolListIndex = 0   ' because the layout has error at right-left way
        cobTrigLevelText = "200"
        cobTrigL_LevelText = "64"  ' because the layout has error
        gnClkDiv = 5000
        
   Else
        cobTrigModeListIndex = 0
        cobTrigSrcListIndex = 1
        cobTrigPolListIndex = 1
        cobClkSrcListIndex = 0
        cobTrigPolListIndex = 1
        cobTrigLevelText = "200"
        cobTrigL_LevelText = "128"
        gnClkDiv = 5000
   End If
   
   '================================
   
    If Left(TmpChip, 10) = "AU3150KLF2" Then
    
        ChipName = TmpChip
    End If
   
     If ChipName = "AU3150KLF23" Or ChipName = "AU3150KLF24" Then
     
     Call MsecDelay(0.5)   'for unstable chip
     End If
     
       If ChipName = "AU3152ALF23" Then
     
     Call MsecDelay(0.5)   'for unstable chip
     End If
   '================================
   
   
   
  
   If Card2214Exist = 0 Then
  Card2214 = D2K_Register_Card(DAQ_2214, 0)
'  Card2214 = D2K_Register_Card(DAQ_2213, 0)
    End If
  ' only set 1 channel
 ' Tester.Cls
  result = D2K_AI_CH_Config(Card2214, ch_num, ch_range Or (ch_ref * 256))
 
 '  Tester.Print "result1="; result
  result = D2K_AI_Config(Card2214, 0, (CLng(cobTrigModeListIndex) * 8) Or CLng(cobTrigSrcListIndex) Or (CLng(cobTrigPolListIndex) * 4096), PostCnt, 0, 1, 1)
  

  
'   Tester.Print "result2="; result
  result = D2K_AI_ContBufferSetup(Card2214, gnBuffer(0), dma_size * ch_cnt, Id)
  
  
 '  Tester.Print "result3="; result
  result = D2K_AIO_Config(Card2214, cobClkSrcListIndex, (CLng(cobTrigPolListIndex) * 256) Or CH0ATRIG, CLng(cobTrigLevelText), CLng(cobTrigL_LevelText))
 
  Status = 0
  '  Tester.Print "result4="; result
  result = D2K_AI_ContReadMultiChannels(Card2214, ch_cnt, ch_num, Id, dma_size, gnClkDiv * ch_cnt, gnClkDiv, ASYNCH_OP)
     
     Tester.Print "result5="; result
  Card2214Exist = 1
  
  
  For i = 0 To 99
  gnBuffer(i) = 0
  Next i
  
  k = 0
  bStop = 0
  OldTimer = Timer
  While Status = 0 And bStop = 0 And (Timer - OldTimer < 10)
  DoEvents
    result = D2K_AI_AsyncCheck(Card2214, Status, count1)
    
  Wend
  result = D2K_AI_AsyncClear(Card2214, startPos, count1)
  
  
 
    
    
     '==========================================================================
    
  ErrCount = 40
  CoutinueCount = 40
 
  If ErrCount > 30 Or CoutinueCount > 30 Then
    If ChipName = "AU3150KLF24" Then    '=========== second mode
       ErrCount = 0
        CoutinueCount = 0
       
       For i = 0 To 99
        Debug.Print MP3Data_21(i), gnBuffer(i)
          If CSng(gnBuffer(i)) = 0 Then
           gnBuffer(i) = 1
           End If
          
        If CSng(MP3Data_21(i)) / CSng(gnBuffer(i)) < 0.7 Or CSng(MP3Data_21(i)) / CSng(gnBuffer(i)) > 1.3 Then
         ErrCount = ErrCount + 1
         If i = OldCount + 1 Then
            CoutinueCount = CoutinueCount + 1
          End If
        End If
        
        OldCount = i
       Next i
        RecordModeMP3 = 21
        
        
        Tester.Print "AU3150BKLF 1 slot pind mode ErrCount="; ErrCount, "64 pin mode CoutinueCount="; CoutinueCount
       
    End If
 End If
    '===========================================================================================
    
   If ErrCount > 30 Or CoutinueCount > 30 Then
    If ChipName = "AU3150KLF24" Then     '=========== second mode
       ErrCount = 0
        CoutinueCount = 0
       
       For i = 0 To 99
        Debug.Print MP3Data_22(i), gnBuffer(i)
          If CSng(gnBuffer(i)) = 0 Then
           gnBuffer(i) = 1
           End If
          
        If CSng(MP3Data_22(i)) / CSng(gnBuffer(i)) < 0.7 Or CSng(MP3Data_22(i)) / CSng(gnBuffer(i)) > 1.3 Then
         ErrCount = ErrCount + 1
         If i = OldCount + 1 Then
            CoutinueCount = CoutinueCount + 1
          End If
        End If
        
        OldCount = i
       Next i
        RecordModeMP3 = 22
        
        
        Tester.Print "AU3150BKLF 1 slot pind mode ErrCount="; ErrCount, "64 pin mode CoutinueCount="; CoutinueCount
       
    End If
 End If
   
     '===========================================================================================
    
   If ErrCount > 30 Or CoutinueCount > 30 Then
    If ChipName = "AU3150KLF24" Then   '=========== second mode
       ErrCount = 0
        CoutinueCount = 0
       
       For i = 0 To 99
        Debug.Print MP3Data_23(i), gnBuffer(i)
          If CSng(gnBuffer(i)) = 0 Then
           gnBuffer(i) = 1
           End If
          
        If CSng(MP3Data_23(i)) / CSng(gnBuffer(i)) < 0.7 Or CSng(MP3Data_23(i)) / CSng(gnBuffer(i)) > 1.3 Then
         ErrCount = ErrCount + 1
         If i = OldCount + 1 Then
            CoutinueCount = CoutinueCount + 1
          End If
        End If
        
        OldCount = i
       Next i
        RecordModeMP3 = 23
       
        
        Tester.Print "AU3150BKLF 1 slot pind mode ErrCount="; ErrCount, "64 pin mode CoutinueCount="; CoutinueCount
       
    End If
 End If
   
   
   
   
  If ErrCount > 30 Or CoutinueCount > 30 Then  ' test criteria
  
  
   AU3130TestMP3 = 2
 
   
   Exit Function
End If
  
  AU3130TestMP3 = 1
 
  
 ' txtXfer.Caption = "Total xfer: " & CStr(count1)
  'scrollView.value = 0
 ' scrollView.Max = Fix((count1 - 1) / DISP_COUNT)
  ' ShowData 0, 20
 ' btnStart.Enabled = True
End Function
Public Function AU3130Test() As Integer
On Error Resume Next
  Dim result As Long
  Dim Status As Byte
  Dim i As Long, k As Long
  Dim ReTrgCnt As Integer, MCnt As Integer, PostCnt As Integer
  Dim TimerSrc As Integer
  Dim cobTrigModeListIndex As Integer
  Dim cobTrigSrcListIndex As Integer
  Dim cobTrigPolListIndex As Integer
  Dim cobClkSrcListIndex As Integer
  Dim ErrCount As Integer
  Dim OldCount As Integer
  Dim CoutinueCount As Integer
  Dim cobTrigLevelText As String
  Dim cobTrigL_LevelText As String
  Dim gnClkDiv As Integer
  Dim Id As Integer
  Dim startPos As Long
  Dim OldTimer
  
   ' dma2205.Hide
  PostCnt = 500 ' post trigger 500
  dma_size = 2000 ' scan count 2000-- 80000
  ch_cnt = 1
  ch_num = 0  ' channel 0
  ch_range = 4  ' BiPolar 1.25V
  ch_ref = 1    ' Differentail
  
  If Left(ChipName, 8) = "AU3130CL" Or Left(ChipName, 8) = "AU3150JL" Or ChipName = "AU3150ALF22" Then
  
        cobTrigModeListIndex = 0
        cobTrigSrcListIndex = 1
        cobTrigPolListIndex = 1
        cobClkSrcListIndex = 0
        cobTrigPolListIndex = 0   ' because the layout has error at right-left way
        cobTrigLevelText = "200"
        cobTrigL_LevelText = "64"  ' because the layout has error
        gnClkDiv = 5000
        
   Else
        cobTrigModeListIndex = 0
        cobTrigSrcListIndex = 1
        cobTrigPolListIndex = 1
        cobClkSrcListIndex = 0
        cobTrigPolListIndex = 1
        cobTrigLevelText = "200"
        cobTrigL_LevelText = "128"
        gnClkDiv = 5000
   End If
   
   '================================
   
    If Left(TmpChip, 10) = "AU3150KLF2" Then
    
        ChipName = TmpChip
    End If
   
     If ChipName = "AU3150KLF23" Then
     
     Call MsecDelay(0.5)   'for unstable chip
     End If
     
       If ChipName = "AU3152ALF23" Then
     
     Call MsecDelay(0.5)   'for unstable chip
     End If
   '================================
   
   
   
  
   If Card2214Exist = 0 Then
  Card2214 = D2K_Register_Card(DAQ_2214, 0)
'  Card2214 = D2K_Register_Card(DAQ_2213, 0)
    End If
  ' only set 1 channel
 ' Tester.Cls
  result = D2K_AI_CH_Config(Card2214, ch_num, ch_range Or (ch_ref * 256))
 
 '  Tester.Print "result1="; result
  result = D2K_AI_Config(Card2214, 0, (CLng(cobTrigModeListIndex) * 8) Or CLng(cobTrigSrcListIndex) Or (CLng(cobTrigPolListIndex) * 4096), PostCnt, 0, 1, 1)
  

  
'   Tester.Print "result2="; result
  result = D2K_AI_ContBufferSetup(Card2214, gnBuffer(0), dma_size * ch_cnt, Id)
  
  
 '  Tester.Print "result3="; result
  result = D2K_AIO_Config(Card2214, cobClkSrcListIndex, (CLng(cobTrigPolListIndex) * 256) Or CH0ATRIG, CLng(cobTrigLevelText), CLng(cobTrigL_LevelText))
 
  Status = 0
  '  Tester.Print "result4="; result
  result = D2K_AI_ContReadMultiChannels(Card2214, ch_cnt, ch_num, Id, dma_size, gnClkDiv * ch_cnt, gnClkDiv, ASYNCH_OP)
     
     Tester.Print "result5="; result
  Card2214Exist = 1
  
  
  For i = 0 To 99
  gnBuffer(i) = 0
  Next i
  
  k = 0
  bStop = 0
  OldTimer = Timer
  While Status = 0 And bStop = 0 And (Timer - OldTimer < 10)
  DoEvents
    result = D2K_AI_AsyncCheck(Card2214, Status, count1)
    
  Wend
  result = D2K_AI_AsyncClear(Card2214, startPos, count1)
  
  
  For i = 0 To 99
  Debug.Print MP3Data(i), gnBuffer(i)
  
     If CSng(gnBuffer(i)) = 0 Then
           gnBuffer(i) = 1
           End If
  If CSng(gnBuffer(i)) <> 0 Then
    If CSng(MP3Data(i)) / CSng(gnBuffer(i)) < 0.7 Or CSng(MP3Data(i)) / CSng(gnBuffer(i)) > 1.3 Then
        ErrCount = ErrCount + 1
        If i = OldCount + 1 Then
           CoutinueCount = CoutinueCount + 1
         End If
    End If
  End If
  OldCount = i
  Next i
    RecordMode = 0
    dma2205.Show
    dma2205.ShowData 0, 100
    
    Tester.Print "ErrCount="; ErrCount, "CoutinueCount="; CoutinueCount
 If ErrCount > 30 Or CoutinueCount > 30 Then
    If ChipName = "AU3130B" Then '=========== second mode
       ErrCount = 0
        CoutinueCount = 0
       
       For i = 0 To 99
        Debug.Print MP3Data_1(i), gnBuffer(i)
           If CSng(gnBuffer(i)) = 0 Then
           gnBuffer(i) = 1
           End If
          
        If CSng(MP3Data_1(i)) / CSng(gnBuffer(i)) < 0.7 Or CSng(MP3Data_1(i)) / CSng(gnBuffer(i)) > 1.3 Then
         ErrCount = ErrCount + 1
         If i = OldCount + 1 Then
            CoutinueCount = CoutinueCount + 1
          
        End If
        End If
        OldCount = i
       Next i
        RecordMode = 2
        dma2205.Show
        dma2205.ShowData 0, 100
        
        Tester.Print "2nd mode ErrCount="; ErrCount, "2 nd mode CoutinueCount="; CoutinueCount
       
    End If
 End If
    
If ErrCount > 30 Or CoutinueCount > 30 Then
    If ChipName = "AU3130B" Then '=========== second mode
       ErrCount = 0
        CoutinueCount = 0
       
       For i = 0 To 99
        Debug.Print MP3Data_1(i), gnBuffer(i)
           If CSng(gnBuffer(i)) = 0 Then
           gnBuffer(i) = 1
           End If
         
        If CSng(MP3Data_2(i)) / CSng(gnBuffer(i)) < 0.7 Or CSng(MP3Data_2(i)) / CSng(gnBuffer(i)) > 1.3 Then
         ErrCount = ErrCount + 1
         If i = OldCount + 1 Then
            CoutinueCount = CoutinueCount + 1
         
        End If
        End If
        OldCount = i
       Next i
        RecordMode = 3
        dma2205.Show
        dma2205.ShowData 0, 100
        
        Tester.Print "3rd mode ErrCount="; ErrCount, "3 rd mode CoutinueCount="; CoutinueCount
       
    End If
 End If
    
    
 If ErrCount > 30 Or CoutinueCount > 30 Then
    If ChipName = "AU3130CL" Then '=========== second mode
       ErrCount = 0
        CoutinueCount = 0
       
       For i = 0 To 99
        Debug.Print MP3Data_3(i), gnBuffer(i)
          If CSng(gnBuffer(i)) = 0 Then
           gnBuffer(i) = 1
           End If
        If CSng(MP3Data_3(i)) / CSng(gnBuffer(i)) < 0.7 Or CSng(MP3Data_3(i)) / CSng(gnBuffer(i)) > 1.3 Then
         ErrCount = ErrCount + 1
         If i = OldCount + 1 Then
            CoutinueCount = CoutinueCount + 1
          End If
        End If
       
        OldCount = i
       Next i
        RecordMode = 4
        dma2205.Show
        dma2205.ShowData 0, 100
        
        Tester.Print "64 1pind mode ErrCount="; ErrCount, "64 pin mode CoutinueCount="; CoutinueCount
       
    End If
 End If
       
    
    
 If ErrCount > 30 Or CoutinueCount > 30 Then
    If ChipName = "AU3130CL" Then '=========== second mode
       ErrCount = 0
        CoutinueCount = 0
       
       For i = 0 To 99
        Debug.Print MP3Data_4(i), gnBuffer(i)
          If CSng(gnBuffer(i)) = 0 Then
           gnBuffer(i) = 1
           End If
          
        If CSng(MP3Data_4(i)) / CSng(gnBuffer(i)) < 0.7 Or CSng(MP3Data_4(i)) / CSng(gnBuffer(i)) > 1.3 Then
         ErrCount = ErrCount + 1
         If i = OldCount + 1 Then
            CoutinueCount = CoutinueCount + 1
          End If
        End If
        
        OldCount = i
       Next i
        RecordMode = 5
        dma2205.Show
        dma2205.ShowData 0, 100
        
        Tester.Print "642 pind mode ErrCount="; ErrCount, "64 pin mode CoutinueCount="; CoutinueCount
       
    End If
 End If
 
 
 
 If ErrCount > 30 Or CoutinueCount > 30 Then
    If ChipName = "AU3130CL" Then '=========== second mode
       ErrCount = 0
        CoutinueCount = 0
       
       For i = 0 To 99
        Debug.Print MP3Data_5(i), gnBuffer(i)
          If CSng(gnBuffer(i)) = 0 Then
           gnBuffer(i) = 1
           End If
          
        If CSng(MP3Data_5(i)) / CSng(gnBuffer(i)) < 0.7 Or CSng(MP3Data_5(i)) / CSng(gnBuffer(i)) > 1.3 Then
         ErrCount = ErrCount + 1
         If i = OldCount + 1 Then
            CoutinueCount = CoutinueCount + 1
          End If
        End If
        
        OldCount = i
       Next i
        RecordMode = 6
        dma2205.Show
        dma2205.ShowData 0, 100
        
        Tester.Print "643 pind mode ErrCount="; ErrCount, "64 pin mode CoutinueCount="; CoutinueCount
       
    End If
 End If
           
           
 If ErrCount > 30 Or CoutinueCount > 30 Then
    If ChipName = "AU3130BLF20" Then '=========== second mode
       ErrCount = 0
        CoutinueCount = 0
       
       For i = 0 To 99
        Debug.Print MP3Data_6(i), gnBuffer(i)
          If CSng(gnBuffer(i)) = 0 Then
           gnBuffer(i) = 1
           End If
          
        If CSng(MP3Data_6(i)) / CSng(gnBuffer(i)) < 0.7 Or CSng(MP3Data_6(i)) / CSng(gnBuffer(i)) > 1.3 Then
         ErrCount = ErrCount + 1
         If i = OldCount + 1 Then
            CoutinueCount = CoutinueCount + 1
          End If
        End If
        
        OldCount = i
       Next i
        RecordMode = 7
        dma2205.Show
        dma2205.ShowData 0, 100
        
        Tester.Print "B44 48 pind mode ErrCount="; ErrCount, "64 pin mode CoutinueCount="; CoutinueCount
       
    End If
 End If
           
 If ErrCount > 30 Or CoutinueCount > 30 Then
    If ChipName = "AU3130BLF20" Then '=========== second mode
       ErrCount = 0
        CoutinueCount = 0
       
       For i = 0 To 99
        Debug.Print MP3Data_7(i), gnBuffer(i)
          If CSng(gnBuffer(i)) = 0 Then
           gnBuffer(i) = 1
           End If
          
        If CSng(MP3Data_7(i)) / CSng(gnBuffer(i)) < 0.7 Or CSng(MP3Data_7(i)) / CSng(gnBuffer(i)) > 1.3 Then
         ErrCount = ErrCount + 1
         If i = OldCount + 1 Then
            CoutinueCount = CoutinueCount + 1
          End If
        End If
        
        OldCount = i
       Next i
        RecordMode = 8
        dma2205.Show
        dma2205.ShowData 0, 100
        
        Tester.Print "B44 48 pind mode ErrCount="; ErrCount, "64 pin mode CoutinueCount="; CoutinueCount
       
    End If
 End If
                      
    
 If ErrCount > 30 Or CoutinueCount > 30 Then
    If ChipName = "AU3130CLF20" Then '=========== second mode
       ErrCount = 0
        CoutinueCount = 0
       
       For i = 0 To 99
        Debug.Print MP3Data_8(i), gnBuffer(i)
          If CSng(gnBuffer(i)) = 0 Then
           gnBuffer(i) = 1
           End If
          
        If CSng(MP3Data_8(i)) / CSng(gnBuffer(i)) < 0.7 Or CSng(MP3Data_8(i)) / CSng(gnBuffer(i)) > 1.3 Then
         ErrCount = ErrCount + 1
         If i = OldCount + 1 Then
            CoutinueCount = CoutinueCount + 1
          End If
        End If
        
        OldCount = i
       Next i
        RecordMode = 9
        dma2205.Show
        dma2205.ShowData 0, 100
        
        Tester.Print "C44 48 pind mode ErrCount="; ErrCount, "64 pin mode CoutinueCount="; CoutinueCount
       
    End If
 End If
                          
    
    
    
    
  If ErrCount > 30 Or CoutinueCount > 30 Then
    If ChipName = "AU3130CLF20" Then '=========== second mode
       ErrCount = 0
        CoutinueCount = 0
       
       For i = 0 To 99
        Debug.Print MP3Data_9(i), gnBuffer(i)
          If CSng(gnBuffer(i)) = 0 Then
           gnBuffer(i) = 1
           End If
          
        If CSng(MP3Data_9(i)) / CSng(gnBuffer(i)) < 0.7 Or CSng(MP3Data_9(i)) / CSng(gnBuffer(i)) > 1.3 Then
         ErrCount = ErrCount + 1
         If i = OldCount + 1 Then
            CoutinueCount = CoutinueCount + 1
          End If
        End If
        
        OldCount = i
       Next i
        RecordMode = 10
        dma2205.Show
        dma2205.ShowData 0, 100
        
        Tester.Print "C44 48 pind mode ErrCount="; ErrCount, "64 pin mode CoutinueCount="; CoutinueCount
       
    End If
 End If
  
'=========================================================================
   If ErrCount > 30 Or CoutinueCount > 30 Then
    If ChipName = "AU3150JLF21" Then '=========== second mode
       ErrCount = 0
        CoutinueCount = 0
       
       For i = 0 To 99
        Debug.Print MP3Data_10(i), gnBuffer(i)
          If CSng(gnBuffer(i)) = 0 Then
           gnBuffer(i) = 1
           End If
          
        If CSng(MP3Data_10(i)) / CSng(gnBuffer(i)) < 0.7 Or CSng(MP3Data_10(i)) / CSng(gnBuffer(i)) > 1.3 Then
         ErrCount = ErrCount + 1
         If i = OldCount + 1 Then
            CoutinueCount = CoutinueCount + 1
          End If
        End If
        
        OldCount = i
       Next i
        RecordMode = 11
        dma2205.Show
        dma2205.ShowData 0, 100
        
        Tester.Print "C44 48 pind mode ErrCount="; ErrCount, "64 pin mode CoutinueCount="; CoutinueCount
       
    End If
 End If
    
'===========================================================================================
    
    
   If ErrCount > 30 Or CoutinueCount > 30 Then
    If ChipName = "AU3150JLF21" Then '=========== second mode
       ErrCount = 0
        CoutinueCount = 0
       
       For i = 0 To 99
        Debug.Print MP3Data_11(i), gnBuffer(i)
          If CSng(gnBuffer(i)) = 0 Then
           gnBuffer(i) = 1
           End If
          
        If CSng(MP3Data_11(i)) / CSng(gnBuffer(i)) < 0.7 Or CSng(MP3Data_11(i)) / CSng(gnBuffer(i)) > 1.3 Then
         ErrCount = ErrCount + 1
         If i = OldCount + 1 Then
            CoutinueCount = CoutinueCount + 1
          End If
        End If
        
        OldCount = i
       Next i
        RecordMode = 12
        dma2205.Show
        dma2205.ShowData 0, 100
        
        Tester.Print "C44 48 pind mode ErrCount="; ErrCount, "64 pin mode CoutinueCount="; CoutinueCount
       
    End If
 End If
       

'===========================================================================================
    
    If ErrCount > 30 Or CoutinueCount > 30 Then
    If ChipName = "AU6254XLS40" Then '=========== second mode
       ErrCount = 0
        CoutinueCount = 0
       
       For i = 0 To 99
        Debug.Print MP3Data_12(i), gnBuffer(i)
          If CSng(gnBuffer(i)) = 0 Then
           gnBuffer(i) = 1
           End If
          
        If CSng(MP3Data_12(i)) / CSng(gnBuffer(i)) < 0.7 Or CSng(MP3Data_12(i)) / CSng(gnBuffer(i)) > 1.3 Then
         ErrCount = ErrCount + 1
         If i = OldCount + 1 Then
            CoutinueCount = CoutinueCount + 1
          End If
        End If
        
        OldCount = i
       Next i
        RecordMode = 13
        dma2205.Show
        dma2205.ShowData 0, 100
        
        Tester.Print "AU6254 pind mode ErrCount="; ErrCount, "64 pin mode CoutinueCount="; CoutinueCount
       
    End If
 End If
    
    
  '===========================================================================================
    
    If ErrCount > 30 Or CoutinueCount > 30 Then
    If ChipName = "AU6254XLS40" Then '=========== second mode
       ErrCount = 0
        CoutinueCount = 0
       
       For i = 0 To 99
        Debug.Print MP3Data_13(i), gnBuffer(i)
          If CSng(gnBuffer(i)) = 0 Then
           gnBuffer(i) = 1
           End If
          
        If CSng(MP3Data_13(i)) / CSng(gnBuffer(i)) < 0.7 Or CSng(MP3Data_13(i)) / CSng(gnBuffer(i)) > 1.3 Then
         ErrCount = ErrCount + 1
         If i = OldCount + 1 Then
            CoutinueCount = CoutinueCount + 1
          End If
        End If
        
        OldCount = i
       Next i
        RecordMode = 14
        dma2205.Show
        dma2205.ShowData 0, 100
        
        Tester.Print "AU6254 pind mode ErrCount="; ErrCount, "64 pin mode CoutinueCount="; CoutinueCount
       
    End If
 End If
    
    
'===========================================================================================
    
    
   If ErrCount > 30 Or CoutinueCount > 30 Then
    If ChipName = "AU3152ALF21" Or ChipName = "AU3152ALF23" Then '=========== second mode
       ErrCount = 0
        CoutinueCount = 0
       
       For i = 0 To 99
        Debug.Print MP3Data_15(i), gnBuffer(i)
          If CSng(gnBuffer(i)) = 0 Then
           gnBuffer(i) = 1
           End If
          
        If CSng(MP3Data_15(i)) / CSng(gnBuffer(i)) < 0.7 Or CSng(MP3Data_15(i)) / CSng(gnBuffer(i)) > 1.3 Then
         ErrCount = ErrCount + 1
         If i = OldCount + 1 Then
            CoutinueCount = CoutinueCount + 1
          End If
        End If
        
        OldCount = i
       Next i
        RecordMode = 15
        dma2205.Show
        dma2205.ShowData 0, 100
        
        Tester.Print "AU3150A52 pind mode ErrCount="; ErrCount, "64 pin mode CoutinueCount="; CoutinueCount
       
    End If
 End If
 '===========================================================================================
  If ErrCount > 30 Or CoutinueCount > 30 Then
    If ChipName = "AU3152ALF21" Or ChipName = "AU3152ALF23" Then '=========== second mode
       ErrCount = 0
        CoutinueCount = 0
       
       For i = 0 To 99
        Debug.Print MP3Data_16(i), gnBuffer(i)
          If CSng(gnBuffer(i)) = 0 Then
           gnBuffer(i) = 1
           End If
          
        If CSng(MP3Data_16(i)) / CSng(gnBuffer(i)) < 0.7 Or CSng(MP3Data_16(i)) / CSng(gnBuffer(i)) > 1.3 Then
         ErrCount = ErrCount + 1
         If i = OldCount + 1 Then
            CoutinueCount = CoutinueCount + 1
          End If
        End If
        
        OldCount = i
       Next i
        RecordMode = 16
        dma2205.Show
        dma2205.ShowData 0, 100
        
        Tester.Print "AU3150A52 pind mode ErrCount="; ErrCount, "64 pin mode CoutinueCount="; CoutinueCount
       
    End If
 End If
 
  '===========================================================================================
  If ErrCount > 30 Or CoutinueCount > 30 Then
    If ChipName = "AU3152ALF21" Or ChipName = "AU3152ALF23" Then '=========== second mode
       ErrCount = 0
        CoutinueCount = 0
       
       For i = 0 To 99
        Debug.Print MP3Data_161(i), gnBuffer(i)
          If CSng(gnBuffer(i)) = 0 Then
           gnBuffer(i) = 1
           End If
          
        If CSng(MP3Data_161(i)) / CSng(gnBuffer(i)) < 0.7 Or CSng(MP3Data_161(i)) / CSng(gnBuffer(i)) > 1.3 Then
         ErrCount = ErrCount + 1
         If i = OldCount + 1 Then
            CoutinueCount = CoutinueCount + 1
          End If
        End If
        
        OldCount = i
       Next i
        RecordMode = 161
        dma2205.Show
        dma2205.ShowData 0, 100
        
        Tester.Print "AU3150A52 pind mode ErrCount="; ErrCount, "64 pin mode CoutinueCount="; CoutinueCount
       
    End If
 End If
 
 
 
 
    '==========================================================================
    
    
'===========================================================================================
    
    
   If ErrCount > 30 Or CoutinueCount > 30 Then
    If ChipName = "AU3150ALF22" Then '=========== second mode
       ErrCount = 0
        CoutinueCount = 0
       
       For i = 0 To 99
        Debug.Print MP3Data_17(i), gnBuffer(i)
          If CSng(gnBuffer(i)) = 0 Then
           gnBuffer(i) = 1
           End If
          
        If CSng(MP3Data_17(i)) / CSng(gnBuffer(i)) < 0.7 Or CSng(MP3Data_17(i)) / CSng(gnBuffer(i)) > 1.3 Then
         ErrCount = ErrCount + 1
         If i = OldCount + 1 Then
            CoutinueCount = CoutinueCount + 1
          End If
        End If
        
        OldCount = i
       Next i
        RecordMode = 17
        dma2205.Show
        dma2205.ShowData 0, 100
        
        Tester.Print "AU3150 2 slot pind mode ErrCount="; ErrCount, "64 pin mode CoutinueCount="; CoutinueCount
       
    End If
 End If
 '===========================================================================================
  If ErrCount > 30 Or CoutinueCount > 30 Then
    If ChipName = "AU3150ALF22" Then '=========== second mode
       ErrCount = 0
        CoutinueCount = 0
       
       For i = 0 To 99
        Debug.Print MP3Data_18(i), gnBuffer(i)
          If CSng(gnBuffer(i)) = 0 Then
           gnBuffer(i) = 1
           End If
          
        If CSng(MP3Data_18(i)) / CSng(gnBuffer(i)) < 0.7 Or CSng(MP3Data_18(i)) / CSng(gnBuffer(i)) > 1.3 Then
         ErrCount = ErrCount + 1
         If i = OldCount + 1 Then
            CoutinueCount = CoutinueCount + 1
          End If
        End If
        
        OldCount = i
       Next i
        RecordMode = 18
        dma2205.Show
        dma2205.ShowData 0, 100
        
        Tester.Print "AU3150 2 slot pind mode ErrCount="; ErrCount, "64 pin mode CoutinueCount="; CoutinueCount
       
    End If
 End If
    '==========================================================================
    
 
  If ErrCount > 30 Or CoutinueCount > 30 Then
    If ChipName = "AU3150KLF21" Then '=========== second mode
       ErrCount = 0
        CoutinueCount = 0
       
       For i = 0 To 99
        Debug.Print MP3Data_19(i), gnBuffer(i)
          If CSng(gnBuffer(i)) = 0 Then
           gnBuffer(i) = 1
           End If
          
        If CSng(MP3Data_19(i)) / CSng(gnBuffer(i)) < 0.7 Or CSng(MP3Data_19(i)) / CSng(gnBuffer(i)) > 1.3 Then
         ErrCount = ErrCount + 1
         If i = OldCount + 1 Then
            CoutinueCount = CoutinueCount + 1
          End If
        End If
        
        OldCount = i
       Next i
        RecordMode = 19
        dma2205.Show
        dma2205.ShowData 0, 100
        
        Tester.Print "AU3150AKLF 1 slot pind mode ErrCount="; ErrCount, "64 pin mode CoutinueCount="; CoutinueCount
       
    End If
 End If
    '===========================================================================================
    
   If ErrCount > 30 Or CoutinueCount > 30 Then
    If ChipName = "AU3150KLF21" Then '=========== second mode
       ErrCount = 0
        CoutinueCount = 0
       
       For i = 0 To 99
        Debug.Print MP3Data_20(i), gnBuffer(i)
          If CSng(gnBuffer(i)) = 0 Then
           gnBuffer(i) = 1
           End If
          
        If CSng(MP3Data_20(i)) / CSng(gnBuffer(i)) < 0.7 Or CSng(MP3Data_20(i)) / CSng(gnBuffer(i)) > 1.3 Then
         ErrCount = ErrCount + 1
         If i = OldCount + 1 Then
            CoutinueCount = CoutinueCount + 1
          End If
        End If
        
        OldCount = i
       Next i
        RecordMode = 20
        dma2205.Show
        dma2205.ShowData 0, 100
        
        Tester.Print "AU3150AKLF 1 slot pind mode ErrCount="; ErrCount, "64 pin mode CoutinueCount="; CoutinueCount
       
    End If
 End If
    
    
     '==========================================================================
    
 
  If ErrCount > 30 Or CoutinueCount > 30 Then
    If ChipName = "AU3150KLF22" Or ChipName = "AU3150KLF23" Then  '=========== second mode
       ErrCount = 0
        CoutinueCount = 0
       
       For i = 0 To 99
        Debug.Print MP3Data_21(i), gnBuffer(i)
          If CSng(gnBuffer(i)) = 0 Then
           gnBuffer(i) = 1
           End If
          
        If CSng(MP3Data_21(i)) / CSng(gnBuffer(i)) < 0.7 Or CSng(MP3Data_21(i)) / CSng(gnBuffer(i)) > 1.3 Then
         ErrCount = ErrCount + 1
         If i = OldCount + 1 Then
            CoutinueCount = CoutinueCount + 1
          End If
        End If
        
        OldCount = i
       Next i
        RecordMode = 21
        dma2205.Show
        dma2205.ShowData 0, 100
        
        Tester.Print "AU3150BKLF 1 slot pind mode ErrCount="; ErrCount, "64 pin mode CoutinueCount="; CoutinueCount
       
    End If
 End If
    '===========================================================================================
    
   If ErrCount > 30 Or CoutinueCount > 30 Then
    If ChipName = "AU3150KLF22" Or ChipName = "AU3150KLF23" Then '=========== second mode
       ErrCount = 0
        CoutinueCount = 0
       
       For i = 0 To 99
        Debug.Print MP3Data_22(i), gnBuffer(i)
          If CSng(gnBuffer(i)) = 0 Then
           gnBuffer(i) = 1
           End If
          
        If CSng(MP3Data_22(i)) / CSng(gnBuffer(i)) < 0.7 Or CSng(MP3Data_22(i)) / CSng(gnBuffer(i)) > 1.3 Then
         ErrCount = ErrCount + 1
         If i = OldCount + 1 Then
            CoutinueCount = CoutinueCount + 1
          End If
        End If
        
        OldCount = i
       Next i
        RecordMode = 22
        dma2205.Show
        dma2205.ShowData 0, 100
        
        Tester.Print "AU3150BKLF 1 slot pind mode ErrCount="; ErrCount, "64 pin mode CoutinueCount="; CoutinueCount
       
    End If
 End If
   
     '===========================================================================================
    
   If ErrCount > 30 Or CoutinueCount > 30 Then
    If ChipName = "AU3150KLF22" Or ChipName = "AU3150KLF23" Then '=========== second mode
       ErrCount = 0
        CoutinueCount = 0
       
       For i = 0 To 99
        Debug.Print MP3Data_23(i), gnBuffer(i)
          If CSng(gnBuffer(i)) = 0 Then
           gnBuffer(i) = 1
           End If
          
        If CSng(MP3Data_23(i)) / CSng(gnBuffer(i)) < 0.7 Or CSng(MP3Data_23(i)) / CSng(gnBuffer(i)) > 1.3 Then
         ErrCount = ErrCount + 1
         If i = OldCount + 1 Then
            CoutinueCount = CoutinueCount + 1
          End If
        End If
        
        OldCount = i
       Next i
        RecordMode = 23
        dma2205.Show
        dma2205.ShowData 0, 100
        
        Tester.Print "AU3150BKLF 1 slot pind mode ErrCount="; ErrCount, "64 pin mode CoutinueCount="; CoutinueCount
       
    End If
 End If
   
   
   
   
  If ErrCount > 30 Or CoutinueCount > 30 Then  ' test criteria
  
  
   AU3130Test = 2
 
   
   Exit Function
End If
  
  AU3130Test = 1
 
  
 ' txtXfer.Caption = "Total xfer: " & CStr(count1)
  'scrollView.value = 0
 ' scrollView.Max = Fix((count1 - 1) / DISP_COUNT)
  ' ShowData 0, 20
 ' btnStart.Enabled = True
End Function

Public Function AU3130TestWMA() As Integer
On Error Resume Next
  Dim result As Long
  Dim Status As Byte
  Dim i As Long, k As Long
  Dim ReTrgCnt As Integer, MCnt As Integer, PostCnt As Integer
  Dim TimerSrc As Integer
  Dim cobTrigModeListIndex As Integer
  Dim cobTrigSrcListIndex As Integer
  Dim cobTrigPolListIndex As Integer
  Dim cobClkSrcListIndex As Integer
  Dim ErrCount As Integer
  Dim OldCount As Integer
  Dim CoutinueCount As Integer
  Dim cobTrigLevelText As String
  Dim cobTrigL_LevelText As String
  Dim gnClkDiv As Integer
  Dim Id As Integer
  Dim startPos As Long
  Dim OldTimer
  
   ' dma2205.Hide
  PostCnt = 500 ' post trigger 500
  dma_size = 2000 ' scan count 2000-- 80000
  ch_cnt = 1
  ch_num = 0  ' channel 0
  ch_range = 4  ' BiPolar 1.25V
  ch_ref = 1    ' Differentail
  
  If Left(ChipName, 8) = "AU3130CL" Or Left(ChipName, 8) = "AU3150JL" Or ChipName = "AU3150ALF22" Then
  
        cobTrigModeListIndex = 0
        cobTrigSrcListIndex = 1
        cobTrigPolListIndex = 1
        cobClkSrcListIndex = 0
        cobTrigPolListIndex = 0   ' because the layout has error at right-left way
        cobTrigLevelText = "200"
        cobTrigL_LevelText = "64"  ' because the layout has error
        gnClkDiv = 5000
        
   Else
        cobTrigModeListIndex = 0
        cobTrigSrcListIndex = 1
        cobTrigPolListIndex = 1
        cobClkSrcListIndex = 0
        cobTrigPolListIndex = 1
        cobTrigLevelText = "200"
        cobTrigL_LevelText = "128"
        gnClkDiv = 5000
   End If
   
   '================================
   
    If Left(TmpChip, 10) = "AU3150KLF2" Then
    
        ChipName = TmpChip
    End If
   
     If ChipName = "AU3150KLF23" Or ChipName = "AU3150KLF24" Then
     
     Call MsecDelay(0.5)   'for unstable chip
     End If
     
       If ChipName = "AU3152ALF23" Then
     
     Call MsecDelay(0.5)   'for unstable chip
     End If
   '================================
   
   
   
  
   If Card2214Exist = 0 Then
  Card2214 = D2K_Register_Card(DAQ_2214, 0)
'  Card2214 = D2K_Register_Card(DAQ_2213, 0)
    End If
  ' only set 1 channel
 ' Tester.Cls
  result = D2K_AI_CH_Config(Card2214, ch_num, ch_range Or (ch_ref * 256))
 
 '  Tester.Print "result1="; result
  result = D2K_AI_Config(Card2214, 0, (CLng(cobTrigModeListIndex) * 8) Or CLng(cobTrigSrcListIndex) Or (CLng(cobTrigPolListIndex) * 4096), PostCnt, 0, 1, 1)
  

  
'   Tester.Print "result2="; result
  result = D2K_AI_ContBufferSetup(Card2214, gnBufferWMA(0), dma_size * ch_cnt, Id)
  
  
 '  Tester.Print "result3="; result
  result = D2K_AIO_Config(Card2214, cobClkSrcListIndex, (CLng(cobTrigPolListIndex) * 256) Or CH0ATRIG, CLng(cobTrigLevelText), CLng(cobTrigL_LevelText))
 
  Status = 0
  '  Tester.Print "result4="; result
  result = D2K_AI_ContReadMultiChannels(Card2214, ch_cnt, ch_num, Id, dma_size, gnClkDiv * ch_cnt, gnClkDiv, ASYNCH_OP)
     
     Tester.Print "result5="; result
  Card2214Exist = 1
  
  
  For i = 0 To 99
  gnBufferWMA(i) = 0   ' need another buffer for repaint at dma show
  Next i
  
  k = 0
  bStop = 0
  OldTimer = Timer
  While Status = 0 And bStop = 0 And (Timer - OldTimer < 10)
  DoEvents
    result = D2K_AI_AsyncCheck(Card2214, Status, count1)
    
  Wend
  result = D2K_AI_AsyncClear(Card2214, startPos, count1)
  
  
  For i = 0 To 99
  Debug.Print MP3WMA_23(i), gnBufferWMA(i)
  
     If CSng(gnBufferWMA(i)) = 0 Then
           gnBufferWMA(i) = 1
           End If
  If CSng(gnBufferWMA(i)) <> 0 Then
    If CSng(MP3WMA_23(i)) / CSng(gnBufferWMA(i)) < 0.7 Or CSng(MP3WMA_23(i)) / CSng(gnBufferWMA(i)) > 1.3 Then
        ErrCount = ErrCount + 1
        If i = OldCount + 1 Then
           CoutinueCount = CoutinueCount + 1
         End If
    End If
  End If
  OldCount = i
  Next i
    RecordModeWMA = 24
   
    
    Tester.Print "ErrCount="; ErrCount, "CoutinueCount="; CoutinueCount
    
    
'===================================================================================
  
    
    If ErrCount > 30 Or CoutinueCount > 30 Then
     If ChipName = "AU3150KLF24" Then   '=========== second mode
        ErrCount = 0
         CoutinueCount = 0
       
        For i = 0 To 99
         Debug.Print MP3WMA_231(i), gnBufferWMA(i)
           If CSng(gnBufferWMA(i)) = 0 Then
            gnBufferWMA(i) = 1
            End If
          
         If CSng(MP3WMA_231(i)) / CSng(gnBufferWMA(i)) < 0.7 Or CSng(MP3WMA_231(i)) / CSng(gnBufferWMA(i)) > 1.3 Then
          ErrCount = ErrCount + 1
          If i = OldCount + 1 Then
             CoutinueCount = CoutinueCount + 1
           End If
         End If
        
         OldCount = i
        Next i
         RecordModeWMA = 241
        
        
         Tester.Print "AU3150BKLF 1 slot pind mode ErrCount="; ErrCount, "64 pin mode CoutinueCount="; CoutinueCount
       
     End If
  End If
   
 '=================================================================================
    If ErrCount > 30 Or CoutinueCount > 30 Then
     If ChipName = "AU3150KLF24" Then   '=========== second mode
        ErrCount = 0
         CoutinueCount = 0
       
        For i = 0 To 99
         Debug.Print MP3WMA_232(i), gnBufferWMA(i)
           If CSng(gnBufferWMA(i)) = 0 Then
            gnBufferWMA(i) = 1
            End If
          
         If CSng(MP3WMA_232(i)) / CSng(gnBufferWMA(i)) < 0.7 Or CSng(MP3WMA_232(i)) / CSng(gnBufferWMA(i)) > 1.3 Then
          ErrCount = ErrCount + 1
          If i = OldCount + 1 Then
             CoutinueCount = CoutinueCount + 1
           End If
         End If
        
         OldCount = i
        Next i
         RecordModeWMA = 242
        
        
         Tester.Print "AU3150BKLF 1 slot pind mode ErrCount="; ErrCount, "64 pin mode CoutinueCount="; CoutinueCount
       
     End If
  End If
    
 '================================================================================
   
   
   
   
  If ErrCount > 30 Or CoutinueCount > 30 Then  ' test criteria
  
  
   AU3130TestWMA = 2
 
   
   Exit Function
End If
  
  AU3130TestWMA = 1
 
  
 ' txtXfer.Caption = "Total xfer: " & CStr(count1)
  'scrollView.value = 0
 ' scrollView.Max = Fix((count1 - 1) / DISP_COUNT)
  ' ShowData 0, 20
 ' btnStart.Enabled = True
End Function


