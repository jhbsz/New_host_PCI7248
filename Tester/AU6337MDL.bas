Attribute VB_Name = "AU6337MDL"
Option Explicit
Public Function InquiryAsciiString(Lun As Byte, CBWDataTransferLength As Long) As Byte
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

'Dim Capacity(0 To 7) As Byte

 
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

 
CBW(8) = &H24  '00
CBW(9) = &H0  '08
CBW(10) = &H0 '00
CBW(11) = &H0 '00

'///////////////  CBW Flag
CBW(12) = &H80                 '80

'////////////// LUN
CBW(13) = Lun                    '00

'///////////// CBD Len
CBW(14) = &H6               '0a

'////////////  UFI command

CBW(15) = &H12
CBW(16) = Lun * 32
 
CBW(17) = &H0         '00
CBW(18) = &H0        '00
CBW(19) = &H24       '00
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
 InquiryAsciiString = 0
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
  InquiryAsciiString = 0
 Exit Function
End If




' ============ Record function
 
                For i = 8 To CBWDataTransferLength - 1
                Debug.Print "k", i, (ReadData(i))
                  If ReadData(i) <> 32 Then
                     If ReadData(i) < 46 Or ReadData(i) > 57 Then
                        If ReadData(i) < 65 Or ReadData(i) > 90 Then
                          If ReadData(i) < 97 Or ReadData(i) > 122 Then
                          
                     
                              InquiryAsciiString = 2  ' card format capacity has problem
                           Exit Function
                          End If
                          End If
                          End If
                          End If
                
                
                Next i

                 
 

 
'3. CSW data
 result = ReadFile _
         (ReadHandle, _
          CSW(0), _
          13, _
          NumberOfBytesRead, _
          HIDOverlapped)  'in
 
If result = 0 Then
 InquiryAsciiString = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
     InquiryAsciiString = 0
Else
     InquiryAsciiString = 1
   
End If

 
End Function
Public Sub AU6337CFF21TestSub()


               Call PowerSet2(1, "5.0", "0.5", 1, "5.0", "0.5", 1)
              '   ChipName = "AU6337BSF20"
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
                
               
                
               
                
          
                    If PCI7248InitFinish = 0 Then
                      PCI7248Exist
                    End If
                
                   result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
           
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                    
                  
                       
                         CardResult = DO_WritePort(card, Channel_P1A, &H3E)  ' 1111 1110
                        
           
              
                     Call MsecDelay(1.4)
      
                    ClosePipe
                    rv0 = CBWTest_New_no_card(0, 1, "vid_058f")
                'Print "a1"
                   Call LabelMenu(0, rv0, 1)
                   ClosePipe
                   Tester.Print rv0; " no card test"
                    
                   
                  If rv0 = 1 Then  ' no card test fail
                   
                   
                        CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' 1111 1110
                        Call MsecDelay(0.02)
      
                      
                   End If
                
    
                
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                '  R/W test
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                
                'initial return value
                
                        ClosePipe
                     rv1 = CBWTest_New(0, 1, "vid_058f")
                    
                     ClosePipe
                    Tester.Print rv1, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                    
                    Tester.Print "LBA="; LBA
                    
                    
                        
                    
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)  ' 1111 1110
                     
               
                     
              
                   If rv0 * rv1 = 1 Then
                         
                            If LightOff <> 254 Then
                               rv1 = 3
                            End If
                            
                            Call LabelMenu(1, rv1, rv0)
                            If rv1 <> 1 Then
                               Tester.Label9.Caption = "SD card Fail or GPO FAIL"
                            End If
                            
              End If
                
                
               
                LBA = LBA + 1
                
                
       '================ INQUIRY SORTING
       
                     If rv0 * rv1 <> 1 Then
                       GoTo AU6337BSLabel
                     End If
                
                    Call PowerSet2(0, "5.3", "0.5", 1, "5.3", "0.5", 1)
                   Call MsecDelay(1.6)
                       rv6 = OpenDriver("vid_058f")
                       
                       rv6 = TestUnitSpeed(0)
                       If rv6 = 1 Then
                         rv6 = TestUnitReady(Lun)
'If TmpInteger = 0 Then
'    TmpInteger = RequestSense(Lun)
    
    
 '   If TmpInteger = 0 Or RequestSenseData(12) <> 58 Then
    
 '      CBWTest_New_no_card = 2  'Write fail
 '      Exit Function
 '   End If
    
'End If
                       rv6 = InquiryAsciiString(0, 36)
                       If rv6 = 1 Then
                       
                         rv6 = InquiryAsciiString(1, 36)
                       Else
                         rv6 = 2
                       End If
                      Else
                        rv6 = 3
                      End If
                      
                     Tester.Print rv6, " \\rv6 :  1 pass ,else inquiry Fail"
                                                 
               ClosePipe
                     Call LabelMenu(2, rv6, rv1)
AU6337BSLabel:                 If rv6 = 2 Then
                     TestResult = "XD_RF"
                     Tester.Label9.Caption = "inquiry fail"
               
                
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
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv1 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                 ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "CF_RF"
                    
                ElseIf rv0 * rv1 = PASS Then
                   TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                End If
                
              '   Call PowerSet2(2, "5.3", "0.5", 1, "5.3", "0.5", 1)
                   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++=

End Sub

Public Sub AU6337BSF21TestSub()


               Call PowerSet2(1, "5.0", "0.5", 1, "5.0", "0.5", 1)
              '   ChipName = "AU6337BSF20"
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
                
               
                
               
                
          
                    If PCI7248InitFinish = 0 Then
                      PCI7248Exist
                    End If
                
                   result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
           
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                    
                  
                       
                         CardResult = DO_WritePort(card, Channel_P1A, &H3E)  ' 1111 1110
                        
           
              
                     Call MsecDelay(1.4)
      
                    ClosePipe
                    rv0 = CBWTest_New_no_card(0, 1, "vid_058f")
                'Print "a1"
                   Call LabelMenu(0, rv0, 1)
                   ClosePipe
                   Tester.Print rv0; " no card test"
                    
                   
                  If rv0 = 1 Then  ' no card test fail
                   
                   
                        CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' 1111 1110
                        Call MsecDelay(0.02)
      
                      
                   End If
                
    
                
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                '  R/W test
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                
                'initial return value
                
                        ClosePipe
                     rv1 = CBWTest_New(0, 1, "vid_058f")
                    
                     ClosePipe
                    Tester.Print rv1, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                    
                    Tester.Print "LBA="; LBA
                    
                    
                        
                    
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)  ' 1111 1110
                     
               
                     
              
                   If rv0 * rv1 = 1 Then
                         
                            If LightOff <> 254 Then
                               rv1 = 3
                            End If
                            
                            Call LabelMenu(1, rv1, rv0)
                            If rv1 <> 1 Then
                               Tester.Label9.Caption = "SD card Fail or GPO FAIL"
                            End If
                            
              End If
                
                
               
                LBA = LBA + 1
                
                
       '================ INQUIRY SORTING
       
                     If rv0 * rv1 <> 1 Then
                       GoTo AU6337BSLabel
                     End If
                
                    Call PowerSet2(0, "5.3", "0.5", 1, "5.3", "0.5", 1)
                   Call MsecDelay(1.6)
                       rv6 = OpenDriver("vid_058f")
                       
                       rv6 = TestUnitSpeed(0)
                       If rv6 = 1 Then
                         rv6 = TestUnitReady(Lun)
'If TmpInteger = 0 Then
'    TmpInteger = RequestSense(Lun)
    
    
 '   If TmpInteger = 0 Or RequestSenseData(12) <> 58 Then
    
 '      CBWTest_New_no_card = 2  'Write fail
 '      Exit Function
 '   End If
    
'End If
                       rv6 = InquiryAsciiString(0, 36)
                       If rv6 = 1 Then
                       
                         rv6 = InquiryAsciiString(1, 36)
                       Else
                         rv6 = 2
                       End If
                      Else
                        rv6 = 3
                      End If
                      
                     Tester.Print rv6, " \\rv6 :  1 pass ,else inquiry Fail"
                                                 
               ClosePipe
                     Call LabelMenu(2, rv6, rv1)
                     
AU6337BSLabel:      If rv6 = 2 Then
                     TestResult = "XD_RF"
                     Tester.Label9.Caption = "inquiry fail"
               
                
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
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv1 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                 ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "CF_RF"
                    
                ElseIf rv0 * rv1 = PASS Then
                   TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                End If
                
              '   Call PowerSet2(2, "5.3", "0.5", 1, "5.3", "0.5", 1)
                   
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++=

End Sub
Public Sub AU6337TestSub()


If ChipName = "AU6337BLF20" Or ChipName = "AU6337GLF20" Or ChipName = "AU6337ILF20" Then
   Call AU6337BLF20TestSub
End If

If ChipName = "AU6337BLF21" Then
    Call AU6337BLF21TestSub
End If

If ChipName = "AU6337BSF21" Then
    Call AU6337BSF21TestSub
End If

If ChipName = "AU6337CFF21" Or ChipName = "AU6337CSF21" Then
    Call AU6337CFF21TestSub
End If

If ChipName = "AU6337ILF21" Then
    Call AU6337ILF21TestSub
End If

If ChipName = "AU6337GLF21" Then
    Call AU6337GLF21TestSub
End If
End Sub

Public Sub AU6337BLF21TestSub()


                Call PowerSet2(1, "5.0", "0.5", 1, "5.0", "0.5", 1)
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
          
                    If PCI7248InitFinish = 0 Then
                      PCI7248Exist
                    End If
                    '   CardResult = DO_WritePort(card, Channel_P1A, &HF)
                       Call MsecDelay(0.2)
                 '  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
                 '   CardResult = DO_WritePort(card, Channel_P1B, &HFF)
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                                                               
                    CardResult = DO_WritePort(card, Channel_P1A, &H6)  ' 0000 0110
        
   
              
                     Call MsecDelay(1.4)
      
                    ClosePipe
                    rv0 = CBWTest_New_no_card(0, 1, "vid_058f")
               '      If rv0 = 1 Then
               '         rv6 = InquiryAsciiString(0, 36)
               '         If rv6 <> 1 Then
               '         rv6 = 2
               '         End If
               '       End If
                    
                     
                   Call LabelMenu(0, rv0, 1)
                   ClosePipe
                     Tester.Print rv0; "SD  no card test"
                    rv1 = CBWTest_New_no_card(1, rv0, "vid_058f")
                    ClosePipe
                    
    
                 
                      
                   Call LabelMenu(3, rv1, rv0)
                    
                 
                   
                    Tester.Print rv1; "MS   no card test"
                   
                
                
                 If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv1 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv1 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
              
                ElseIf rv1 * rv0 = PASS Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
                    
                   
                  If rv0 = 1 And rv1 = 1 Then   ' no card test fail
                     rv0 = 0
                     rv1 = 0
                     
               
         
                     
                        CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' 1111 1110
                        Call MsecDelay(0.02)
                        
                       ClosePipe
                     rv0 = CBWTest_New(0, 1, "vid_058f")
                    
                     ClosePipe
                      Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                          
                      Call LabelMenu(1, rv0, 1)
                     rv1 = CBWTest_New(1, rv0, "vid_058f")
                    
                     ClosePipe
                      Tester.Print rv1, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                                                 
         
                    
                   Tester.Print "LBA="; LBA
                    
                    
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)  ' 1111 1110
                       
                                         
                            If rv0 * rv1 = 1 Then
                        
                            
                                If LightOff <> 252 Then
                                   rv3 = 3
                                    Tester.Label9.Caption = " GPO FAIL"
                                
                                End If
                            End If
                            
                            
                                If Left(ChipName, 8) = "AU6337GL" Then
                        
                            
                                If LightOff <> 254 Then
                                   rv3 = 3
                                    Tester.Label9.Caption = " GPO FAIL"
                                
                                End If
                            End If
                            Call LabelMenu(3, rv1, rv0)
                           
          
            
                LBA = LBA + 1
                
                
                
                
                   If rv0 * rv1 <> 1 Then
                    GoTo AU6337BLLabel
                   End If
                
                   Call PowerSet2(0, "5.3", "0.5", 1, "5.3", "0.5", 1)
                   Call MsecDelay(1.6)
                       rv6 = OpenDriver("vid_058f")
                       
                       rv6 = TestUnitSpeed(0)
                       If rv6 = 1 Then
                         rv6 = TestUnitReady(Lun)
'If TmpInteger = 0 Then
'    TmpInteger = RequestSense(Lun)
    
    
 '   If TmpInteger = 0 Or RequestSenseData(12) <> 58 Then
    
 '      CBWTest_New_no_card = 2  'Write fail
 '      Exit Function
 '   End If
    
'End If
                       rv6 = InquiryAsciiString(0, 36)
                       If rv6 = 1 Then
                       
                         rv6 = InquiryAsciiString(1, 36)
                       Else
                         rv6 = 2
                       End If
                      Else
                        rv6 = 3
                      End If
                      
                    
               ClosePipe
                     Call LabelMenu(2, rv6, rv1)
                     
                 Tester.Print rv6, " \\rv6 :  1 pass ,else inquiry Fail"
                                                 
              
AU6337BLLabel:                 If rv6 = 2 Then
                     TestResult = "XD_RF"
                     Tester.Label9.Caption = "inquiry fail"
                ElseIf rv0 = UNKNOW Then
                    UnknowDeviceFail = UnknowDeviceFail + 1
                    TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Or rv6 = 3 Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                 ElseIf rv1 = WRITE_FAIL Or rv3 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv1 = READ_FAIL Or rv3 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
    
                ElseIf rv0 * rv1 = PASS Then
                  
                        
                        
                Else
                    TestResult = "Bin2"
                End If
                
            End If
               
               
              '    Call PowerSet2(2, "5.3", "0.5", 1, "5.3", "0.5", 1)
                
End Sub
                
                
Public Sub AU6337GLF21TestSub()


                Call PowerSet2(1, "5.0", "0.5", 1, "5.0", "0.5", 1)
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
          
                    If PCI7248InitFinish = 0 Then
                      PCI7248Exist
                    End If
                    '   CardResult = DO_WritePort(card, Channel_P1A, &HF)
                       Call MsecDelay(0.2)
                 '  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
                 '   CardResult = DO_WritePort(card, Channel_P1B, &HFF)
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                                                               
                    CardResult = DO_WritePort(card, Channel_P1A, &H6)  ' 0000 0110
        
   
              
                     Call MsecDelay(1.4)
      
                    ClosePipe
                    rv0 = CBWTest_New_no_card(0, 1, "vid_058f")
               '      If rv0 = 1 Then
               '         rv6 = InquiryAsciiString(0, 36)
               '         If rv6 <> 1 Then
               '         rv6 = 2
               '         End If
               '       End If
                    
                     
                   Call LabelMenu(0, rv0, 1)
                   ClosePipe
                     Tester.Print rv0; "SD  no card test"
                    rv1 = CBWTest_New_no_card(1, rv0, "vid_058f")
                    ClosePipe
                    
    
                 
                      
                   Call LabelMenu(3, rv1, rv0)
                    
                 
                   
                    Tester.Print rv1; "MS   no card test"
                   
                
                
                 If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv1 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv1 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
              
                ElseIf rv1 * rv0 = PASS Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
                    
                   
                  If rv0 = 1 And rv1 = 1 Then   ' no card test fail
                     rv0 = 0
                     rv1 = 0
                     
               
         
                     
                        CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' 1111 1110
                        Call MsecDelay(0.02)
                        
                       ClosePipe
                     rv0 = CBWTest_New(0, 1, "vid_058f")
                    
                     ClosePipe
                      Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                          
                      Call LabelMenu(1, rv0, 1)
                     rv1 = CBWTest_New(1, rv0, "vid_058f")
                    
                     ClosePipe
                      Tester.Print rv1, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                                                 
         
                    
                   Tester.Print "LBA="; LBA
                    
                    
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)  ' 1111 1110
                       
                                         
                            If rv0 * rv1 = 1 Then
                             If LightOff <> 254 Then
                                   rv3 = 3
                                    Tester.Label9.Caption = " GPO FAIL"
                                
                                End If
                            
                               
                            End If
                            
                            
                             
                            
                               
                           
                            Call LabelMenu(3, rv1, rv0)
                           
          
            
                LBA = LBA + 1
                
                
                
                
                   If rv0 * rv1 <> 1 Then
                    GoTo AU6337BLLabel
                   End If
                
                   Call PowerSet2(0, "5.3", "0.5", 1, "5.3", "0.5", 1)
                   Call MsecDelay(1.6)
                       rv6 = OpenDriver("vid_058f")
                       
                       rv6 = TestUnitSpeed(0)
                       If rv6 = 1 Then
                         rv6 = TestUnitReady(Lun)
'If TmpInteger = 0 Then
'    TmpInteger = RequestSense(Lun)
    
    
 '   If TmpInteger = 0 Or RequestSenseData(12) <> 58 Then
    
 '      CBWTest_New_no_card = 2  'Write fail
 '      Exit Function
 '   End If
    
'End If
                       rv6 = InquiryAsciiString(0, 36)
                       If rv6 = 1 Then
                       
                         rv6 = InquiryAsciiString(1, 36)
                       Else
                         rv6 = 2
                       End If
                      Else
                        rv6 = 3
                      End If
                      
                    
               ClosePipe
                     Call LabelMenu(2, rv6, rv1)
                     
                 Tester.Print rv6, " \\rv6 :  1 pass ,else inquiry Fail"
                                                 
              
AU6337BLLabel:                 If rv6 = 2 Then
                     TestResult = "XD_RF"
                     Tester.Label9.Caption = "inquiry fail"
                ElseIf rv0 = UNKNOW Then
                    UnknowDeviceFail = UnknowDeviceFail + 1
                    TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Or rv6 = 3 Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                 ElseIf rv1 = WRITE_FAIL Or rv3 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv1 = READ_FAIL Or rv3 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
    
                ElseIf rv0 * rv1 = PASS Then
                  
                        
                        
                Else
                    TestResult = "Bin2"
                End If
                
            End If
               
               
              '    Call PowerSet2(2, "5.3", "0.5", 1, "5.3", "0.5", 1)
                
End Sub

Public Sub AU6337ILF21TestSub()


                Call PowerSet2(1, "5.0", "0.5", 1, "5.0", "0.5", 1)
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
          
                    If PCI7248InitFinish = 0 Then
                      PCI7248Exist
                    End If
                    '   CardResult = DO_WritePort(card, Channel_P1A, &HF)
                       Call MsecDelay(0.2)
                 '  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
                 '   CardResult = DO_WritePort(card, Channel_P1B, &HFF)
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                                                               
                    CardResult = DO_WritePort(card, Channel_P1A, &H6)  ' 0000 0110
        
   
              
                     Call MsecDelay(1.4)
      
                    ClosePipe
                    rv0 = CBWTest_New_no_card(0, 1, "vid_058f")
               '      If rv0 = 1 Then
               '         rv6 = InquiryAsciiString(0, 36)
               '         If rv6 <> 1 Then
               '         rv6 = 2
               '         End If
               '       End If
                    
                     
                   Call LabelMenu(0, rv0, 1)
                   ClosePipe
                     Tester.Print rv0; "SD  no card test"
                    rv1 = CBWTest_New_no_card(1, rv0, "vid_058f")
                    ClosePipe
                    
    
                 
                      
                   Call LabelMenu(3, rv1, rv0)
                    
                 
                   
                    Tester.Print rv1; "MS   no card test"
                   
                
                
                 If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv1 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv1 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
              
                ElseIf rv1 * rv0 = PASS Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
                    
                   
                  If rv0 = 1 And rv1 = 1 Then   ' no card test fail
                     rv0 = 0
                     rv1 = 0
                     
               
         
                     
                        CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' 1111 1110
                        Call MsecDelay(0.02)
                        
                       ClosePipe
                     rv0 = CBWTest_New(0, 1, "vid_058f")
                    
                     ClosePipe
                      Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                          
                      Call LabelMenu(1, rv0, 1)
                     rv1 = CBWTest_New(1, rv0, "vid_058f")
                    
                     ClosePipe
                      Tester.Print rv1, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                                                 
         
                    
                   Tester.Print "LBA="; LBA
                    
                    
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)  ' 1111 1110
                       
                                         
                            If rv0 * rv1 = 1 Then
                        
                            
                                If LightOff <> 252 Then
                                   rv3 = 3
                                    Tester.Label9.Caption = " GPO FAIL"
                                
                                End If
                            End If
                            
                            
                                If Left(ChipName, 8) = "AU6337GL" Then
                        
                            
                                If LightOff <> 254 Then
                                   rv3 = 3
                                    Tester.Label9.Caption = " GPO FAIL"
                                
                                End If
                            End If
                            Call LabelMenu(3, rv1, rv0)
                           
          
            
                LBA = LBA + 1
                
                
                
                
                   If rv0 * rv1 <> 1 Then
                    GoTo AU6337BLLabel
                   End If
                
                   Call PowerSet2(0, "5.3", "0.5", 1, "5.3", "0.5", 1)
                   Call MsecDelay(1.6)
                       rv6 = OpenDriver("vid_058f")
                       
                       rv6 = TestUnitSpeed(0)
                       If rv6 = 1 Then
                         rv6 = TestUnitReady(Lun)
'If TmpInteger = 0 Then
'    TmpInteger = RequestSense(Lun)
    
    
 '   If TmpInteger = 0 Or RequestSenseData(12) <> 58 Then
    
 '      CBWTest_New_no_card = 2  'Write fail
 '      Exit Function
 '   End If
    
'End If
                       rv6 = InquiryAsciiString(0, 36)
                       If rv6 = 1 Then
                       
                         rv6 = InquiryAsciiString(1, 36)
                       Else
                         rv6 = 2
                       End If
                      Else
                        rv6 = 3
                      End If
                      
                    
               ClosePipe
                     Call LabelMenu(2, rv6, rv1)
                     
                 Tester.Print rv6, " \\rv6 :  1 pass ,else inquiry Fail"
                                                 
              
AU6337BLLabel:                 If rv6 = 2 Then
                     TestResult = "XD_RF"
                     Tester.Label9.Caption = "inquiry fail"
                ElseIf rv0 = UNKNOW Then
                    UnknowDeviceFail = UnknowDeviceFail + 1
                    TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Or rv6 = 3 Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                 ElseIf rv1 = WRITE_FAIL Or rv3 = WRITE_FAIL Then
                    MSWriteFail = MSWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv1 = READ_FAIL Or rv3 = READ_FAIL Then
                    MSReadFail = MSReadFail + 1
                    TestResult = "MS_RF"
    
                ElseIf rv0 * rv1 = PASS Then
                  
                        
                        
                Else
                    TestResult = "Bin2"
                End If
                
            End If
               
               
              '    Call PowerSet2(2, "5.3", "0.5", 1, "5.3", "0.5", 1)
                
End Sub
Public Sub AU6337BLF20TestSub()
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
   
                    If PCI7248InitFinish = 0 Then
                      PCI7248Exist
                    End If
                
                 '  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
                 '   CardResult = DO_WritePort(card, Channel_P1B, &HFF)
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                                                               
                    CardResult = DO_WritePort(card, Channel_P1A, &H6)  ' 0000 0110
        
   
              
                     Call MsecDelay(1.4)
      
                    ClosePipe
                    rv0 = CBWTest_New_no_card(0, 1, "vid_058f")
                  
               
                   Call LabelMenu(0, rv0, 1)
                   ClosePipe
                     Tester.Print rv0; "SD  no card test"
                    rv1 = CBWTest_New_no_card(1, rv0, "vid_058f")
                    
                 
                   Call LabelMenu(3, rv1, rv0)
                   
                    Tester.Print rv1; "MS   no card test"
                   
                   
                         
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv1 = WRITE_FAIL Then
                    CFWriteFail = CFWriteFail + 1
                    TestResult = "CF_WF"
                ElseIf rv1 = READ_FAIL Then
                    CFReadFail = CFReadFail + 1
              
                ElseIf rv1 * rv0 = PASS Then
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                  
                End If
                    
                   
                  If rv0 = 1 And rv1 = 1 Then ' no card test fail
                     rv0 = 0
                     rv1 = 0
                     
               
         
                     
                        CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' 1111 1110
                        Call MsecDelay(0.02)
                        
                       ClosePipe
                     rv0 = CBWTest_New(0, 1, "vid_058f")
                    
                     ClosePipe
                      Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                          
                      Call LabelMenu(1, rv0, 1)
                     rv1 = CBWTest_New(1, rv0, "vid_058f")
                    
                     ClosePipe
                      Tester.Print rv1, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                                                 
         
                    
                   Tester.Print "LBA="; LBA
                    
                    
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)  ' 1111 1110
                       
                                         
                            If Left(ChipName, 8) = "AU6337BL" Or Left(ChipName, 8) = "AU6337IL" Then
                        
                            
                                If LightOff <> 252 Then
                                   rv3 = 3
                                    Tester.Label9.Caption = " GPO FAIL"
                                
                                End If
                            End If
                            
                            
                                If Left(ChipName, 8) = "AU6337GL" Then
                        
                            
                                If LightOff <> 254 Then
                                   rv3 = 3
                                    Tester.Label9.Caption = " GPO FAIL"
                                
                                End If
                            End If
                            Call LabelMenu(3, rv1, rv0)
                           
          
            
                LBA = LBA + 1
               
                
                If rv0 = UNKNOW Then
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
                
                        
                    
                ElseIf rv0 * rv1 = PASS Then
                  
                        
                        
                Else
                    TestResult = "Bin2"
                End If
                
            End If
               
               
                
                CardResult = DO_WritePort(card, Channel_P1A, &H1)
End Sub
