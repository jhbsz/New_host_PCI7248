Attribute VB_Name = "AU6371Mdl"
Option Explicit
Dim LBA1 As Long


Public Sub AU6336ZFF20TestSub()
  
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
                
                
'                    If PCI7248InitFinish = 0 Then
'                      PCI7248Exist
'                    End If
                
                If PCI7248InitFinish_Sync = 0 Then
                    PCI7248Exist_P1C_Sync
                End If
                
                   result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
           
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                    
                    
                    
                       CardResult = DO_WritePort(card, Channel_P1A, &HE)   '
                    
              
                     Call MsecDelay(0.8)
      
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
                     rv1 = CBWTest_NewAU6336ZFF20(0, rv0, "vid_058f")
                    
                     ClosePipe
                    Tester.Print rv1, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                    
                    Tester.Print "LBA="; LBA
                    
                    
                        
                    
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)  ' 1111 1110
                     
               
                     
                       
                      
                          
                            If LightOff <> 254 Then
                               rv1 = 3
                            End If
                            
                            Call LabelMenu(1, rv1, rv0)
                            If rv1 <> 1 Then
                               Tester.Label9.Caption = "SD card Fail or GPO FAIL"
                            End If
                            
                     
                     
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
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv1 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                 ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                    
                ElseIf rv0 * rv1 = PASS Then
                    
                     TestResult = "PASS"
                Else
                    TestResult = "Bin2"
                End If
                
                
                CardResult = DO_WritePort(card, Channel_P1A, &H1)
                
                
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++=

End Sub


Public Sub NotShareBusSingleSlotTest()
  
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
                
                If ChipName = "AU6332GFF20" Or ChipName = "AU6332FFF20" Then
                ChipName = "AU6332BSF0"
                End If
                
                
                If ChipName = "AU6336DFF20" Then
                
                 ChipName = "AU6336AFF20"
                 
               End If
                
         
               If ChipName = "AU6332CF" Or ChipName = "AU6371CF" Or ChipName = "AU6332BSF0" Or ChipName = "AU6337CFF20" Or ChipName = "AU6371DFT10" Or ChipName = "AU6331CSFT10" Or ChipName = "AU6337BSF20" Or ChipName = "AU6337CSF20" Or ChipName = "AU6336AFF20" Then
         
                    If PCI7248InitFinish = 0 Then
                      PCI7248Exist
                    End If
                
                   result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
           
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                    
                    If ChipName = "AU6371CF" Or ChipName = "AU6336AFF20" Then
                       'bit0: ENA
                       'bit1: SDCDN
                       'bin2: Reader mode
                       'bin3:
                       'bin4,5,6,7: DATA
                       
                    
                       CardResult = DO_WritePort(card, Channel_P1A, &HE)   '
                    ElseIf ChipName = "AU6331CSFT10" Then
                         'bit0: ENA
                       'bit1: SDCDN
                       'bin2: Reader mode
                       'bin3:
                       'bin4,5,6,7: DATA
                    
                         CardResult = DO_WritePort(card, Channel_P1A, &HE)   '
                          result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
                       ChipName = "AU6371CF"
                       
                    ElseIf ChipName = "AU6332CF" Or ChipName = "AU6371DFT10" Or ChipName = "AU6337BSF20" Or ChipName = "AU6337CSF20" Or ChipName = "AU6337CFF20" Then
                        CardResult = DO_WritePort(card, Channel_P1A, &H3E)  ' 1111 1110
                        
                        
                    Else 'ChipName = "AU6332BS"
                    CardResult = DO_WritePort(card, Channel_P1A, &H8)   ' 1111 1110
                      result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
           
                     
    
                        
                    End If
                   
              
                     Call MsecDelay(1.4)
      
                    ClosePipe
                    rv0 = CBWTest_New_no_card(0, 1, "vid_058f")
                'Print "a1"
                   Call LabelMenu(0, rv0, 1)
                   ClosePipe
                   Tester.Print rv0; " no card test"
                    
                   
                  If rv0 = 1 Then  ' no card test fail
                   
                      If ChipName = "AU6371CF" Or ChipName = "AU6371DFT10" Then
                         CardResult = DO_WritePort(card, Channel_P1A, &HC)  ' 1111 1110
                        Call MsecDelay(0.02)
                      
                      Else
                        CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' 1111 1110
                        Call MsecDelay(0.02)
                        
                        
                      End If
                      
                      
                   End If
               End If
               
               
               
                
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                '  R/W test
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                
                'initial return value
                
                
                
             
                
                If ChipName = "AU6332CF" Or ChipName = "AU6371CF" Or ChipName = "AU6332BSF0" Or ChipName = "AU6337CFF20" Or ChipName = "AU6371DFT10" Or ChipName = "AU6337BSF20" Or ChipName = "AU6337CSF20" Or ChipName = "AU6336AFF20" Then
                
                      ClosePipe
                     rv1 = CBWTest_New(0, 1, "vid_058f")
                    
                     ClosePipe
                    Tester.Print rv1, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                    
                    Tester.Print "LBA="; LBA
                    
                    
                        
                    
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)  ' 1111 1110
                     
               
                     
                       
                       If ChipName = "AU6371CF" Then
                       
                            If LightOff <> 224 And LightOff <> 254 Then
                               rv1 = 3
                            End If
                            Call LabelMenu(1, rv1, rv0)
                            If rv1 <> 1 And rv0 = 1 Then
                                   Tester.Label9.Caption = "SD card Fail or GPO FAIL  "
                            End If
                              
                            rv2 = 4
                            If rv1 = 1 Then
                            
                                '  PullDown Test
                                '1. shutdown power, Bus switch Enable, unable reader mode
                                
                                CardResult = DO_WritePort(card, Channel_P1A, &H1)
                                Call MsecDelay(0.3)
                                CardResult = DO_WritePort(card, Channel_P1A, &H0)
                                Call MsecDelay(0.3)
                                CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                rv2 = 1
                                If LightOff <> 225 And LightOff <> 224 Then
                                        rv2 = 2
                                End If
                                
                                If rv2 = 1 Then
                                    CardResult = DO_WritePort(card, Channel_P1A, &HF0)
                                    Call MsecDelay(0.3)
                                    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                                    If LightOff <> 255 And LightOff <> 254 Then
                                            rv2 = 2
                                    End If
                                End If
                            End If
                             Call LabelMenu(2, rv2, rv1)
                            If rv2 <> 1 And rv2 <> 4 Then
                                   Tester.Label9.Caption = "TriState Test Fail "
                            End If
                             
                              
                            
                            
                       End If
                       
                       
                         If ChipName = "AU6332CF" Or ChipName = "AU6371DFT10" Then
                            If LightOff <> 252 Then
                               rv1 = 3
                            End If
                            
                            Call LabelMenu(1, rv1, rv0)
                            If rv1 <> 1 Then
                               Tester.Label9.Caption = "SD card Fail or GPO FAIL"
                            End If
                            
                       End If
                       
                       
                         If ChipName = "AU6337BSF20" Or ChipName = "AU6337CSF20" Or ChipName = "AU6336AFF20" Then
                         
                            If LightOff <> 254 Then
                               rv1 = 3
                            End If
                            
                            Call LabelMenu(1, rv1, rv0)
                            If rv1 <> 1 Then
                               Tester.Label9.Caption = "SD card Fail or GPO FAIL"
                            End If
                            
                            
                       End If
                       
                        If ChipName = "AU6332BSF0" Then
                            If LightOff <> 239 Then
                               rv1 = 3
                            End If
                            
                            Call LabelMenu(1, rv1, rv0)
                            If rv1 <> 1 Then
                               Tester.Label9.Caption = "SD card Fail or GPO FAIL"
                            End If
                            
                       End If
                       
                      If ChipName = "AU6337CFF20" Then
                            If LightOff <> 254 Then
                               rv1 = 3
                            End If
                            
                            Call LabelMenu(1, rv1, rv0)
                            If rv1 <> 1 Then
                               Tester.Label9.Caption = "SD card Fail or GPO FAIL"
                            End If
                            
                       End If
                    
                  
                
                
                Else
                
                      ClosePipe
                    rv0 = CBWTest_New(0, 1, "vid_058f")
                    Call LabelMenu(1, rv0, 1)
                    ClosePipe
                     Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                    
                   Tester.Print "LBA="; LBA
                
                  
               
                   
                 
                     
                    
                 
                
                     
                End If
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
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv1 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                 ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                    
                ElseIf rv0 = PASS Then
                    If ChipName = "AU6332CF" Or ChipName = "AU6337BSF20" Then
                      If rv1 = 1 Then
                        TestResult = "PASS"
                      Else
                        TestResult = "Bin2"
                      End If
                    Else
                      TestResult = "PASS"
                    End If
                        
                    If ChipName = "AU6371CF" Then
                      If rv1 * rv2 = 1 Then
                        TestResult = "PASS"
                      Else
                        TestResult = "Bin2"
                      End If
                    Else
                      TestResult = "PASS"
                    End If
                        
                        
                Else
                    TestResult = "Bin2"
                End If
                
                If ChipName = "AU6332CF" Or ChipName = "AU6371CF" Or ChipName = "AU6337BSF20" Then
                
                CardResult = DO_WritePort(card, Channel_P1A, &H1)
                End If
                
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++=

End Sub
Public Sub NotShareBusSingleSlotTestAU6336DFF21TestSub()
Tester.Print "this is Normal mode"
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
                    
                    
                       
                    
                CardResult = DO_WritePort(card, Channel_P1A, &HE)   '
                     
                   ' test card detect
              
                Call MsecDelay(0.8)
                    ClosePipe
                    rv0 = CBWTest_New_no_card(0, 1, "vid_058f")
                
                    Call LabelMenu(0, rv0, 1)
                    ClosePipe
                    Tester.Print rv0; " no card test"
                    
                   
                  If rv0 = 1 Then  ' no card test fail
                        CardResult = DO_ReadPort(card, Channel_P1B, LightOff)  ' 1111 1110
                      
                    
                      ' 1111 1110
                       CardResult = DO_WritePort(card, Channel_P1A, &H0)
                        Call MsecDelay(0.02)
                          
                        
                          ClosePipe
                         rv1 = CBWTest_New(0, 1, "vid_058f")
                          ClosePipe
                     
                      CardResult = DO_ReadPort(card, Channel_P1B, LightOn)  ' 1111 1110
                      
                      Tester.Print rv1, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                    
                      Tester.Print "LBA="; LBA
                        Call LabelMenu(1, rv1, rv0)
                     If rv1 = 1 Then
                         
                           
                        If LightOff <> 255 Or LightOn <> 254 Then
                               rv1 = 3
                                Call LabelMenu(1, rv1, rv0)
                               Tester.Label9.Caption = "SD card Fail or GPO FAIL"
                        End If
             
                     
                       End If
               
                  End If
                  
                  
                  
               
                 
                      
       
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
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv1 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                 ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                    
                ElseIf rv0 * rv1 = PASS Then
                    TestResult = "PASS"
                        
                Else
                    TestResult = "Bin2"
                End If
                
                
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++=

End Sub
 
Public Sub NotShareBusSingleSlotTestAU6336DFF21TestSub_WPTest()
Tester.Print "this is Normal mode"
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
                    
                    
                       
                    
                CardResult = DO_WritePort(card, Channel_P1A, &HE)   '
                     
                   ' test card detect
              
                Call MsecDelay(0.8)
                    ClosePipe
                    rv0 = CBWTest_New_no_card(0, 1, "vid_058f")
                
                    Call LabelMenu(0, rv0, 1)
                    ClosePipe
                    Tester.Print rv0; " no card test"
                    
                   
                  If rv0 = 1 Then  ' no card test fail
                        CardResult = DO_ReadPort(card, Channel_P1B, LightOff)  ' 1111 1110
                      
                    
                      ' 1111 1110
                       CardResult = DO_WritePort(card, Channel_P1A, &H0)
                        Call MsecDelay(0.02)
                          
                        
                          ClosePipe
                         rv1 = CBWTest_New(0, 1, "vid_058f")
                            CardResult = DO_WritePort(card, Channel_P1A, &HE)
                           Call MsecDelay(0.05)
                           CardResult = DO_WritePort(card, Channel_P1A, &H80)
                         Call MsecDelay(0.05)
                      
                     
                      CardResult = DO_ReadPort(card, Channel_P1B, LightOn)  ' 1111 1110
                      
                      Tester.Print rv1, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                    
                      Tester.Print "LBA="; LBA
                        Call LabelMenu(1, rv1, rv0)
                     If rv1 = 1 Then
                         
                           
                        If LightOff <> 255 Or LightOn <> 254 Then
                               rv1 = 3
                                Call LabelMenu(1, rv1, rv0)
                               Tester.Label9.Caption = "SD card Fail or GPO FAIL"
                        End If
             
                     
                       End If
               
                  End If
                  
                  
                  
               
                  If rv0 * rv1 = 1 Then
               
                           rv2 = TestUnitReady(Lun)
                       
                            rv2 = RequestSense(Lun)
                         rv2 = Read_Data1(LBA, 0, 1024)
                        rv2 = Write_Data2(LBA, 0, 1024)
                        Tester.Print "WP :"; rv2
                          rv2 = TestUnitReady(Lun)
                       
                            rv2 = RequestSense(Lun)
                            
                           
                       '     Dim i As Integer
                           ' Tester.Cls
                        '    For i = 0 To 17
                        '    Debug.Print i, Hex(RequestSenseData(i))
                        '    Next
                       
                         If Hex(RequestSenseData(2)) = 7 And Hex(RequestSenseData(12)) = 27 Then
                          rv2 = 1
                         Else
                          rv2 = 2
                            
                        End If
                            
                     ClosePipe
                  End If
                   
                      Call LabelMenu(2, rv2, rv1)
                      
                      If rv2 = 2 Then
                      Tester.Label9.Caption = "write protest test Fail"
                      End If
                      
       
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
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv1 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                 ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                    
                ElseIf rv0 * rv1 * rv2 = PASS Then
                    TestResult = "PASS"
                        
                Else
                    TestResult = "Bin2"
                End If
                
                
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++=

End Sub
  
Public Sub NotShareBusSingleSlotTestAU6336LFF21TestSub_WPTest()
Tester.Print " NB mode test"

Dim ChipString As String
  
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
                    
                  
                       'bit1: SDCDN
                       'bin2: Reader mode
                       'bin3:
                       'bin4,5,6,7: DATA
                       
                       
                        CardResult = DO_WritePort(card, Channel_P1A, &HE)  ' pwr low, cdn high
                        
                       ' nb mode and cdn test
                       Call MsecDelay(0.5)     'power on time
                       ChipString = "vid"
                       If GetDeviceName(ChipString) <> "" Then
                          rv0 = 0
                         GoTo AU6336Result
                  
                       End If
              
            
             
                     CardResult = DO_WritePort(card, Channel_P1A, &HC)   'Cdn low
              
                     Call MsecDelay(1.2)
  
                
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                '  R/W test
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                
                'initial return value
                
  
                     ClosePipe
                     rv0 = CBWTest_New(0, 1, "vid_058f")
                      ClosePipe
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                    Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                    
                    Tester.Print "LBA="; LBA
                    
                      
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)  ' 1111 1110
                       Call LabelMenu(0, rv0, 1)
                           
                     If rv0 = 1 Then
                        'rv2 = Write_Data_LED(LBA, 0, 1024)  ' to get led
                       
                          rv1 = 1
                             If LightOn <> 254 Or LightOff <> 255 Then
                               rv1 = 2
                            End If
   
                            Call LabelMenu(1, rv1, rv0)
                            If rv1 <> 1 Then
                               Tester.Label9.Caption = "   GPO FAIL"
                            End If
                            
                    End If
                    
                       If rv0 * rv1 = 1 Then
                           
               
                            CardResult = DO_WritePort(card, Channel_P1A, &HE)
                           Call MsecDelay(0.4)
                           CardResult = DO_WritePort(card, Channel_P1A, &H80)
                         Call MsecDelay(1.2)
                      
                         OpenPipe
                           rv2 = TestUnitReady(Lun)
                       
                            rv2 = RequestSense(Lun)
                         rv2 = Read_Data1(LBA, 0, 1024)
                        rv2 = Write_Data2(LBA, 0, 1024)
                        Tester.Print "WP :"; rv2
                          rv2 = TestUnitReady(Lun)
                       
                            rv2 = RequestSense(Lun)
                            
                           
                         '    Dim i As Integer
                           ' Tester.Cls
                        '    For i = 0 To 17
                        '    Debug.Print i, Hex(RequestSenseData(i))
                        '    Next
                       
                         If Hex(RequestSenseData(2)) = 7 And Hex(RequestSenseData(12)) = 27 Then
                          rv2 = 1
                         Else
                          rv2 = 2
                            
                        End If
                            
                     ClosePipe
                  End If
                   
                      Call LabelMenu(2, rv2, rv1)
                      
                      If rv2 = 2 Then
                      Tester.Label9.Caption = "write protest test Fail"
                      End If
                    
              
                    
                      ClosePipe
                      
                       
                  
                
             
                LBA = LBA + 1
               
AU6336Result:
                If rv0 = UNKNOW Then
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
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                    
                ElseIf rv0 * rv1 = PASS Then
                     
                       TestResult = "PASS"
                        
                Else
                    TestResult = "Bin2"
                End If
                
           
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++=

End Sub
Public Sub AU6336LFF01TestSub()
'2013/2/21 Use S/B:AU6479GFL 48LQ SOCKET
'2013/4/26 for C63,C64 FT3


Dim TmpLBA As Long
Dim i As Integer
Dim DetectCount As Integer
Dim HV_Done_Flag As Boolean
Dim HV_Result As String
Dim LV_Result As String


If PCI7248InitFinish_Sync = 0 Then
    PCI7248Exist_P1C_Sync
End If

Dim ChipString As String

Routine_Label:

ReaderExist = 0
ChipString = "pid_6335"


If Not HV_Done_Flag Then
    Call PowerSet2(0, "5.25", "0.3", 1, "5.25", "0.3", 0)
    'Call MsecDelay(0.1)
    Tester.Print "AU6336LF : 3.6V Begin Test ..."
    SetSiteStatus (RunHV)
Else
    Call PowerSet2(0, "4.75", "0.3", 1, "4.75", "0.3", 0)
    'Call MsecDelay(0.1)
    Tester.Print vbCrLf & "AU6336LF : 3.0V Begin Test ..."
    SetSiteStatus (RunLV)
End If
DetectCount = 0


OldChipName = ""
               
LBA = LBA + 1
         
rv0 = 0     'Enum
rv1 = 0     'SD (Lun0)

Tester.Label3.BackColor = RGB(255, 255, 255)
Tester.Label4.BackColor = RGB(255, 255, 255)
Tester.Label5.BackColor = RGB(255, 255, 255)
Tester.Label6.BackColor = RGB(255, 255, 255)
Tester.Label7.BackColor = RGB(255, 255, 255)
Tester.Label8.BackColor = RGB(255, 255, 255)


' Ctrl_1            Ctrl_2

'8:                 8:
'7:                 7:
'6:                 6:
'5:                 5:

'4:                 4:
'3:                 3:
'2:SDCDN            2:
'1:ENA            1:GPON7

'=========================================
'    POWER on
'=========================================

CardResult = DO_WritePort(card, Channel_P1A, &HC)      'SDCDN¡BENA On
If CardResult <> 0 Then
    MsgBox "Set SD Card Detect Down Fail"
    End
End If


'===============================================
'  Enum Device
'===============================================
                    
Call MsecDelay(0.3)
rv0 = WaitDevOn(ChipString)
Call MsecDelay(0.2)

Call NewLabelMenu(0, "WaitDevice", rv0, 1)

If rv0 = 1 Then
    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
    If CardResult <> 0 Then
        MsgBox "Read light On fail"
        End
    End If

    Call MsecDelay(0.02)
    
    If ((LightOn And &H1) <> 0) Then
        Tester.Print "LightON="; LightOn
        'Tester.Print "LightOFF="; LightOff
        UsbSpeedTestResult = GPO_FAIL
        rv1 = 3
        Call NewLabelMenu(1, "LED", rv1, rv0)
    End If
End If

If LBA < 1000 Then
    LBA = 1000
Else
    LBA = LBA + 1
End If

Tester.Print "LBA="; LBA
'ClosePipe

'===============================================
'  SD Card test
'===============================================

If rv0 = 1 Then
    'rv1 = CBWTest_New_AU6366LFF21(0, rv0, ChipString)
    rv1 = CBWTest_New_6336(0, rv0, ChipString)
    If rv1 <> 1 Then
        Call NewLabelMenu(1, "SD", rv1, rv0)
    Else
'        rv1 = CBWTest_New_128_Sector_PipeReady(0, rv1)    ' write
'        Call NewLabelMenu(1, "SD_64K", rv1, rv0)
'
'        If rv1 = 1 Then
            rv1 = Read_SD_Speed_AU6371(0, 0, 18, "4Bits")
            
            Call NewLabelMenu(1, "SD Bus Speed/Width", rv1, rv0)
'        End If
        
    End If

    Tester.Print rv1, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
End If

If rv1 = 1 Then
    CardResult = DO_WritePort(card, Channel_P1A, &HE)
    Call MsecDelay(0.2)
    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
    If ((LightOn And &H1) <> 1) Then
        Tester.Print "LightOFF="; LightOn
        UsbSpeedTestResult = GPO_FAIL
        rv1 = 3
        Call NewLabelMenu(1, "LED", rv1, rv0)
    End If
    
    If rv1 = 1 Then
        If GetDeviceName_NoReply(ChipString) <> "" Then
            rv1 = 3
            Call NewLabelMenu(1, "NBMD", rv1, rv0)
        End If
    End If
End If


                 

AU6336LFResult:
                
    ClosePipe
    CardResult = DO_WritePort(card, Channel_P1A, &HE)
    If Not HV_Done_Flag Then
        SetSiteStatus (HVDone)
        Call WaitAnotherSiteDone(HVDone, 4#)
    Else
        SetSiteStatus (LVDone)
        Call WaitAnotherSiteDone(LVDone, 4#)
    End If
    Call PowerSet2(0, "0.0", "0.5", 1, "0.0", "0.5", 1)
    WaitDevOFF (ChipString)
    SetSiteStatus (SiteUnknow)
    
    If HV_Done_Flag = False Then
        If rv0 <> 1 Then
            HV_Result = "Bin2"
            Tester.Print "HV Unknow"
        ElseIf rv0 * rv1 <> 1 Then
            HV_Result = "Fail"
            Tester.Print "HV Fail"
        ElseIf rv0 * rv1 = 1 Then
            HV_Result = "PASS"
            Tester.Print "HV PASS"
        End If
        
        HV_Done_Flag = True
        Call MsecDelay(0.3)
        GoTo Routine_Label
    Else
        If rv0 <> 1 Then
            LV_Result = "Bin2"
            Tester.Print "LV Unknow"
        ElseIf rv0 * rv1 <> 1 Then
            LV_Result = "Fail"
            Tester.Print "LV Fail"
        ElseIf rv0 * rv1 = 1 Then
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

Public Sub NotShareBusSingleSlotTestAU6336LFF21TestSub()
Tester.Print " NB mode test"

Dim ChipString As String
  
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
                    
                  
                       'bit1: SDCDN
                       'bin2: Reader mode
                       'bin3:
                       'bin4,5,6,7: DATA
  
                     CardResult = DO_WritePort(card, Channel_P1A, &HC)   'Cdn low
              
                     Call MsecDelay(0.8)
  
                
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                '  R/W test
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                
                'initial return value
                
  
                     ClosePipe
                     rv0 = CBWTest_New_AU6366LFF21(0, 1, "vid_058f")
                     ClosePipe
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                    Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                    
                    Tester.Print "LBA="; LBA
                    
                      
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)  ' 1111 1110
                       Call LabelMenu(0, rv0, 1)
                           
                     If rv0 = 1 Then
                       ' rv2 = Write_Data_LED(LBA, 0, 1024)  ' to get led
                        
                          rv1 = 1
                             If LightOn <> 254 Or LightOff <> 255 Then
                               rv1 = 2
                            End If
   
                            Call LabelMenu(1, rv1, rv0)
                            If rv1 <> 1 Then
                               Tester.Label9.Caption = "   GPO FAIL"
                            End If
                            
                    End If
                    
                       If rv1 = 1 Then
                          CardResult = DO_WritePort(card, Channel_P1A, &HE)  ' pwr low, cdn high
                          rv2 = 1
                         ' nb mode and cdn test
                          Call MsecDelay(0.4)     'power on time
                          ChipString = "vid"
                          If GetDeviceName(ChipString) <> "" Then
                             rv2 = 2
                             Call LabelMenu(1, rv2, rv1)
                    
                          End If
                 
                       End If
                       
                       
                   CardResult = DO_WritePort(card, Channel_P1A, &HC)   'Cdn low
                
             
                LBA = LBA + 1
               
AU6336Result:
                If rv0 = UNKNOW Then
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
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                    
                ElseIf rv0 * rv1 * rv2 = PASS Then
                     
                       TestResult = "PASS"
                        
                Else
                    TestResult = "Bin2"
                End If
                
           
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++=

End Sub
Public Sub NotShareBusSingleSlotTestAU6336IF()
  
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
                    
                  
                       'bit1: SDCDN
                       'bin2: Reader mode
                       'bin3:
                       'bin4,5,6,7: DATA
                       
                    
                       CardResult = DO_WritePort(card, Channel_P1A, &HC)   '
                 
                     
    
                    
                   
              
                     Call MsecDelay(1.4)
      
       
             
               
                
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                '  R/W test
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                
                'initial return value
                
                
                
             
                
                   
                      ClosePipe
                     rv0 = CBWTest_New(0, 1, "vid_058f")
                    
                     ClosePipe
                    Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                    
                    Tester.Print "LBA="; LBA
                    
                     Call LabelMenu(0, rv0, 1)
                        
                    
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)  ' 1111 1110
                     
               
                 
                       
                      
                       
                         rv1 = 1
                        
                         
                            If LightOff <> 254 Then
                               rv1 = 2
                            End If
                            
                            Call LabelMenu(1, rv1, rv0)
                            If rv1 <> 1 Then
                               Tester.Label9.Caption = "   GPO FAIL"
                            End If
                            
                            
                      
                       
                  
                
             
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
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv1 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                 ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                    
                ElseIf rv0 * rv1 = PASS Then
                     
                       TestResult = "PASS"
                        
                Else
                    TestResult = "Bin2"
                End If
                
           
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++=

End Sub

Public Sub AU6336IFF01TestSub()
  
Dim ChipString As String
Dim TempStr As String
Dim HV_Done_Flag As Boolean
Dim HV_Result As String
Dim LV_Result As String


If PCI7248InitFinish_Sync = 0 Then
    PCI7248Exist_P1C_Sync
End If


Routine_Label:

'result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
'CardResult = DO_WritePort(card, Channel_P1B, &H0)
'
'result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)

Call MsecDelay(0.2)

If Not HV_Done_Flag Then
    Call PowerSet2(0, "3.6", "0.5", 1, "3.6", "0.5", 0)
    Call MsecDelay(0.1)
    Tester.Print "AU6336 : HV Begin Test ..."
    SetSiteStatus (RunHV)
Else
    Call PowerSet2(0, "3.3", "0.5", 1, "3.3", "0.5", 0)
    Call MsecDelay(0.1)
    Tester.Print vbCrLf & "AU6336 : LV Begin Test ..."
    SetSiteStatus (RunLV)
End If

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
                

'result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
'CardResult = DO_WritePort(card, Channel_P1B, &H0)
'
'result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                  
'bit1: SDCDN
'bin2: Reader mode
'bin3:
'bin4,5,6,7: DATA
                       
 
   '
CardResult = DO_WritePort(card, Channel_P1A, &HC)
WaitDevOn ("pid_6335")
Call MsecDelay(0.02)
      
          
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'
'  R/W test
'
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

'initial return value
                
                   
ClosePipe
rv0 = CBWTest_New(0, 1, "pid_6335")
ClosePipe

Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"

Tester.Print "LBA="; LBA

Call LabelMenu(0, rv0, 1)


'result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
CardResult = DO_ReadPort(card, Channel_P1B, LightOff)  ' 1111 1110
                     
rv1 = 1
                        
                         
If LightOff <> 254 Then
    rv1 = 2
End If

Call LabelMenu(1, rv1, rv0)
If rv1 <> 1 Then
    Tester.Label9.Caption = "   GPO FAIL"
End If
                            
               
             
LBA = LBA + 1
                
' nb mode test==========================
                
rv2 = 1
CardResult = DO_WritePort(card, Channel_P1A, &HE)  ' pwr low, cdn high
                        
                        
' nb mode and cdn test
Call MsecDelay(0.2)     'power on time
ChipString = "pid_6335"

If GetDeviceName_NoReply(ChipString) = "" Then
    rv2 = 1
End If

If rv2 = 2 Then
    Tester.Label9.Caption = "NB mode Fail"
End If

Call PowerSet2(0, "0.0", "0.5", 1, "0.0", "0.5", 0)
Call MsecDelay(0.1)
CardResult = DO_WritePort(card, Channel_P1A, &H0)

'CardResult = DIO_PortConfig(card, Channel_P1A, INPUT_PORT)
'Call MsecDelay(0.02)
'CardResult = DIO_PortConfig(card, Channel_P1A, OUTPUT_PORT)

If Not HV_Done_Flag Then
    SetSiteStatus (HVDone)
    Call WaitAnotherSiteDone(HVDone, 3#)
Else
    SetSiteStatus (LVDone)
    Call WaitAnotherSiteDone(LVDone, 3#)
End If


WaitDevOFF ("pid_6335")
SetSiteStatus (SiteUnknow)

If HV_Done_Flag = False Then
    If rv0 <> 1 Then
        HV_Result = "Bin2"
        Tester.Print "HV Unknow"
    ElseIf rv0 * rv1 * rv2 <> 1 Then
        HV_Result = "Fail"
        Tester.Print "HV Fail"
    ElseIf rv0 * rv1 * rv2 = 1 Then
        HV_Result = "PASS"
        Tester.Print "HV PASS"
    End If
    HV_Done_Flag = True
    Call MsecDelay(0.3)
    GoTo Routine_Label
Else
    If rv0 <> 1 Then
        LV_Result = "Bin2"
        Tester.Print "LV Unknow"
    ElseIf rv0 * rv1 * rv2 <> 1 Then
        LV_Result = "Fail"
        Tester.Print "LV Fail"
    ElseIf rv0 * rv1 * rv2 = 1 Then
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

Public Sub NotShareBusSingleSlotTestAU6336IFF21TestSub()
  
Dim ChipString As String
  
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
                    
                  
                       'bit1: SDCDN
                       'bin2: Reader mode
                       'bin3:
                       'bin4,5,6,7: DATA
                       
                    
                       CardResult = DO_WritePort(card, Channel_P1A, &HC)   '
                 
                     
    
                    
                   
              
                     Call MsecDelay(1#)
      
       
             
               
                
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                '  R/W test
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                
                'initial return value
                
                
                
             
                
                   
                      ClosePipe
                     rv0 = CBWTest_New(0, 1, "vid_058f")
                    
                     ClosePipe
                    Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                    
                    Tester.Print "LBA="; LBA
                    
                     Call LabelMenu(0, rv0, 1)
                        
                    
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)  ' 1111 1110
                     
               
                 
                       
                      
                       
                         rv1 = 1
                        
                         
                            If LightOff <> 254 Then
                               rv1 = 2
                            End If
                            
                            Call LabelMenu(1, rv1, rv0)
                            If rv1 <> 1 Then
                               Tester.Label9.Caption = "   GPO FAIL"
                            End If
                            
                            
                      
                       
                  
                
             
                LBA = LBA + 1
                
                ' nb mode test==========================
                
                 rv2 = 1
                        CardResult = DO_WritePort(card, Channel_P1A, &HE)  ' pwr low, cdn high
                        
                       ' nb mode and cdn test
                       Call MsecDelay(0.2)     'power on time
                       ChipString = "vid"
                       If GetDeviceName(ChipString) <> "" Then
                          rv2 = 2
                     
                       End If
               
                      CardResult = DO_WritePort(card, Channel_P1A, &HC)
                Call LabelMenu(2, rv2, rv1)
                
                If rv2 = 2 Then
                Tester.Label9.Caption = "NB mode Fail"
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
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv1 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                 ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                    
                ElseIf rv0 * rv1 * rv2 = PASS Then
                     
                       TestResult = "PASS"
                        
                Else
                    TestResult = "Bin2"
                End If
                
           
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++=

End Sub


Public Sub NotShareBusSingleSlotTestAU6336HFF21TestSub()
  
Dim ChipString As String
  
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
                    
                  
                       'bit1: SDCDN
                       'bin2: Reader mode
                       'bin3:
                       'bin4,5,6,7: DATA
                       
                    
                       CardResult = DO_WritePort(card, Channel_P1A, &HC)   '
                 
                     
    
                    
                   
              
                     Call MsecDelay(1#)
      
       
             
               
                
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                '  R/W test
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                
                'initial return value
                
                
                
             
                
                   
                      ClosePipe
                     rv0 = CBWTest_New(0, 1, "vid_058f")
                    
                     ClosePipe
                    Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                    
                    Tester.Print "LBA="; LBA
                    
                     Call LabelMenu(0, rv0, 1)
                        
                    
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)  ' 1111 1110
                     
               
                 
                       
                      
                       
                         rv1 = 1
                        
                         
                            If LightOff <> 254 Then
                               rv1 = 2
                            End If
                            
                            Call LabelMenu(1, rv1, rv0)
                            If rv1 <> 1 Then
                               Tester.Label9.Caption = "   GPO FAIL"
                            End If
                            
                            
                      
                       
                  
                
             
                LBA = LBA + 1
                
                ' nb mode test==========================
                
                 rv2 = 1
                        CardResult = DO_WritePort(card, Channel_P1A, &HE)  ' pwr low, cdn high
                        
                       ' nb mode and cdn test
                       Call MsecDelay(0.2)     'power on time
                       ChipString = "vid"
                       If GetDeviceName(ChipString) = "" Then
                          rv2 = 2
                     
                       End If
               
                      CardResult = DO_WritePort(card, Channel_P1A, &HC)
                Call LabelMenu(2, rv2, rv1)
                
                If rv2 = 2 Then
                Tester.Label9.Caption = "Normal mode Fail"
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
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv1 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                 ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                    
                ElseIf rv0 * rv1 * rv2 = PASS Then
                     
                       TestResult = "PASS"
                        
                Else
                    TestResult = "Bin2"
                End If
                
           
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++=

End Sub

Public Sub NotShareBusSingleSlotTestAU6336CAF20()
  
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
                    
                  
                       'bit1: SDCDN
                       'bin2: Reader mode
                       'bin3:
                       'bin4,5,6,7: DATA
                       
                    
                       CardResult = DO_WritePort(card, Channel_P1A, &HD)   '
                 
                     
    
                    
                   
              
                     Call MsecDelay(1#)
      
       
                     rv0 = NBModeTestFcn("058f")
               
                
                    CardResult = DO_WritePort(card, Channel_P1A, &HC)   '
                 
           
                     Call MsecDelay(1#)
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                '  R/W test
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                
                'initial return value
                
                
                
             
                
                   
                      ClosePipe
                     rv0 = CBWTest_New(0, rv0, "vid")
                    
                     ClosePipe
                    Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                    
                    Tester.Print "LBA="; LBA
                    
                     Call LabelMenu(0, rv0, 1)
                        
                    
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)  ' 1111 1110
                     
               
                 
                       
                      
                       
                         rv1 = 1
                        
                         
                            If LightOff <> 254 Then
                               rv1 = 2
                            End If
                            
                            Call LabelMenu(1, rv1, rv0)
                            If rv1 <> 1 Then
                               Tester.Label9.Caption = "   GPO FAIL"
                            End If
                            
                            
                      
                       
                  
                
             
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
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv1 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                 ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                    
                ElseIf rv0 * rv1 = PASS Then
                     
                       TestResult = "PASS"
                        
                Else
                    TestResult = "Bin2"
                End If
                
           
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++=

End Sub
Public Sub NotShareBusSingleSlotTestAU6336ExtRom()
  
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
                    
                  
                       'bit1: SDCDN
                       'bin2: Reader mode
                       'bin3:
                       'bin4,5,6,7: DATA
                       
                    
                       CardResult = DO_WritePort(card, Channel_P1A, &HC)   '
                 
                     
    
                    
                   
              
                     Call MsecDelay(1.4)
      
       
             
               
                
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '
                '  R/W test
                '
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                
                'initial return value
                
                
                
             
                
                   
                      ClosePipe
                     rv0 = CBWTest_New(0, 1, "vid_1984")
                    
                     ClosePipe
                    Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                    
                    Tester.Print "LBA="; LBA
                    
                     Call LabelMenu(0, rv0, 1)
                        
                    
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)  ' 1111 1110
                     
               
                 
                       
                      
                       
                         rv1 = 1
                        
                         
                            If LightOff <> 254 Then
                               rv1 = 2
                            End If
                            
                            Call LabelMenu(1, rv1, rv0)
                            If rv1 <> 1 Then
                               Tester.Label9.Caption = "   GPO FAIL"
                            End If
                            
                            
                      
                       
                  
                
             
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
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv1 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                 ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                    
                ElseIf rv0 * rv1 = PASS Then
                     
                       TestResult = "PASS"
                        
                Else
                    TestResult = "Bin2"
                End If
                
           
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++=

End Sub
Public Function AU6986Physical_Write_Data(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte

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
CBW(13) = 0                   '00

'///////////// CBD Len
CBW(14) = &HA                '0a

'////////////  UFI command


'////////////  UFI command

CBW(15) = &HFA  ' vendor out command
CBW(16) = &H2   ' physical write
 
CBW(17) = 4 ' sector len
CBW(18) = 0  'zone

LBAByte(0) = (LBA1 Mod 256)
tmpV(0) = Int(LBA1 / 256)
LBAByte(1) = (tmpV(0) Mod 256)


CBW(19) = LBAByte(1)   ' sector addr
CBW(20) = LBAByte(0)


CBW(21) = 0   ' sector addr
CBW(22) = 0


 
CBW(23) = LBAByte(1)     ' logical addr
CBW(24) = LBAByte(0)
For i = 25 To 30
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
    AU6986Physical_Write_Data = 0
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
    AU6986Physical_Write_Data = 0
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
    AU6986Physical_Write_Data = 0
    Exit Function
End If
 
 
 
If CSW(12) = 1 Then
AU6986Physical_Write_Data = 0

Else
AU6986Physical_Write_Data = 1
End If
End Function
Function AU6986Physical_Read_Data(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte
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

CBW(8) = 0  '00
CBW(9) = &H6    '08
CBW(10) = 0 '00
CBW(11) = 0 '00

'///////////////  CBW Flag
CBW(12) = &H80                 '80

'////////////// LUN
CBW(13) = 0                    '00

'///////////// CBD Len
CBW(14) = &H8             '0a


'////////////  UFI command

CBW(15) = &HFA  ' vendor out command
CBW(16) = &H1  ' physical read
 
CBW(17) = 4 ' sector len
CBW(18) = 0  'zone
'CBW(19) = &H0  'block addr
'CBW(20) = &H80

'CBW(19) = &H1  'block addr
'CBW(20) = &HF6

 
LBAByte(0) = (LBA1 Mod 256)
tmpV(0) = Int(LBA1 / 256)
LBAByte(1) = (tmpV(0) Mod 256)

CBW(19) = LBAByte(1)   ' sector addr
CBW(20) = LBAByte(0)

CBW(21) = 0 ' sector addr
CBW(22) = 0
 
 
For i = 23 To 30
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
 AU6986Physical_Read_Data = 0
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
 AU6986Physical_Read_Data = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
    AU6986Physical_Read_Data = 0
Else
     AU6986Physical_Read_Data = 1
   
End If

  For i = 0 To 2047
    
        If ReadData(i) <> Pattern(i) Then
          AU6986Physical_Read_Data = 3   'Read fail
          Exit Function
        End If
    
    Next
  
End Function

Function AU6986Physical_Erase(LBA As Long, Lun As Byte, CBWDataTransferLength As Long) As Byte
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
CBW(12) = &H0                 '80

'////////////// LUN
CBW(13) = 0                    '00

'///////////// CBD Len
CBW(14) = &H5                '0a

'////////////  UFI command

CBW(15) = &HFA  ' vendor out command
CBW(16) = &H3  ' physical read
 
 
CBW(17) = 0  'zone

LBAByte(0) = (LBA1 Mod 256)
tmpV(0) = Int(LBA1 / 256)
LBAByte(1) = (tmpV(0) Mod 256)

CBW(18) = LBAByte(1)  ' vendor out command
CBW(19) = LBAByte(0) ' physical read
 


 
 
For i = 20 To 30
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
 AU6986Physical_Erase = 0
 Exit Function
End If

'2. Readdata stage
 
'result = ReadFile _
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
 AU6986Physical_Erase = 0
 Exit Function
End If
 
'4. CSW status

If CSW(12) = 1 Then
    AU6986Physical_Erase = 0
Else
     AU6986Physical_Erase = 1
   
End If

 
End Function
Public Sub AU6986TestSub()

'this routine is for AU6986 add physical R/W test


              If Left(ChipName, 8) = "AU6986HL" Then
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
                   End If
                
                If LBA1 > 950 Then
                LBA1 = 0
                End If
                
                
          
                
                
                LBA = LBA + 1
                LBA1 = LBA + 500
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                Call LabelMenu(0, rv0, 1)
                ClosePipe
               
               If rv0 = 1 Then
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If LightOn <> 223 Then
                        rv0 = 2
                         Call LabelMenu(0, rv0, 1)
                          Tester.Label9.Caption = "GPO Fail"
                     End If
               End If
               
               
               
                 
               If rv0 = 1 Then
                 OpenPipe
                  rv1 = AU6986Physical_Erase(0, 0, 2048)
                 ClosePipe
               End If
                
               
                
               If rv1 = 1 Then
                 OpenPipe
                   rv2 = AU6986Physical_Write_Data(0, 0, 2048)
                  ClosePipe
               End If
               
                If rv2 = 1 Then
                    OpenPipe
                    rv3 = AU6986Physical_Read_Data(0, 0, 2560)
                    ClosePipe
                    
                End If
        
        
        
                Tester.Print rv0, " \\ : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " physical erase : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " physical write : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " physical read : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                 Tester.Print "LBA="; LBA
        
              rv2 = rv1 * rv2 * rv3
              
              If rv2 <> 1 Then
                If rv1 <> 1 Then
                rv2 = 4
                Else
                rv2 = 2
                End If
              End If
                
                
                
              
             Call LabelMenu(2, rv2, rv0)
               
               
            
            
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                ElseIf rv0 * rv2 = PASS Then
                     TestResult = "PASS"
                   
                Else
                      TestResult = "Bin2"
                End If



End Sub
Public Sub AU6986ALTestSub()

'this routine is for AU6986 add physical R/W test


              If Left(ChipName, 8) = "AU6986AL" Then
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
                   End If
                
                If LBA1 > 950 Then
                LBA1 = 0
                End If
                
                
               
                
                
                LBA = LBA + 1
                LBA1 = LBA + 500
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                 Call LabelMenu(0, rv0, 1)
                ClosePipe
               
               If rv0 = 1 Then
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                     If LightOn <> 254 Then
                        rv0 = 2
                         Call LabelMenu(0, rv0, 1)
                          Tester.Label9.Caption = "GPO Fail"
                     End If
               End If
               
              
               
                 
               If rv0 = 1 Then
                 OpenPipe
                  rv1 = AU6986Physical_Erase(0, 0, 2048)
                 ClosePipe
               End If
                
               
                
               If rv1 = 1 Then
                 OpenPipe
                   rv2 = AU6986Physical_Write_Data(0, 0, 2048)
                  ClosePipe
               End If
               
                If rv2 = 1 Then
                    OpenPipe
                    rv3 = AU6986Physical_Read_Data(0, 0, 2560)
                    ClosePipe
                    
                End If
        
        
        
                Tester.Print rv0, " \\ : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv1, " physical erase : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv2, " physical write : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                Tester.Print rv3, " physical read : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                 Tester.Print "LBA="; LBA
        
              rv2 = rv1 * rv2 * rv3
              
              If rv2 <> 1 Then
                If rv1 <> 1 Then
                rv2 = 4
                Else
                rv2 = 2
                End If
              End If
                
                
                
              
             Call LabelMenu(2, rv2, rv0)
               
               
            
            
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv2 = WRITE_FAIL Then
                    XDWriteFail = XDWriteFail + 1
                    TestResult = "XD_WF"
                ElseIf rv2 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                    TestResult = "XD_RF"
                ElseIf rv0 * rv2 = PASS Then
                     TestResult = "PASS"
                   
                Else
                      TestResult = "Bin2"
                End If



End Sub

Public Sub AU6395BLTestSub()
     
                    If PCI7248InitFinish = 0 Then
                      PCI7248Exist
                    End If
                
                   result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    
                    CardResult = DO_WritePort(card, Channel_P1A, &H80)  ' 1111 1110
                    Call MsecDelay(0.1)
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                    CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' 1111 1110
                    Call MsecDelay(0.8)
                 
                
                  
                
                
                LBA = LBA + 1
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                Call LabelMenu(0, rv0, 1)
                ClosePipe
               ' If rv0 = 1 Then
               '     TestResult = "PASS"
               ' End If
               
                If rv0 = 1 Then
                     rv1 = 1
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                     If LightOn <> 31 Then
                        rv1 = 2
                     End If
                 End If
               
                Tester.Print rv0, " \\ : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                Tester.Print rv1, " \\ : 1 GPO pass ,2 : GPO fail"
                   Call LabelMenu(2, rv1, rv0)
                   If rv1 = 2 Then
                   Tester.Label9.Caption = "GPO Fail"
                   End If
                
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
                XDWriteFail = XDWriteFail + 1
                TestResult = "XD_WF"
                ElseIf rv1 = READ_FAIL Then
                XDReadFail = XDReadFail + 1
                TestResult = "XD_RF"
                    
                    
                ElseIf rv0 * rv1 = PASS Then
                     TestResult = "PASS"
                     
                Else
                      TestResult = "Bin2"
                End If
End Sub

Public Sub AU6395CLTestSub()
     
                    If PCI7248InitFinish = 0 Then
                      PCI7248Exist
                    End If
                
                   result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    
                    CardResult = DO_WritePort(card, Channel_P1A, &H80)  ' 1111 1110
                    Call MsecDelay(0.1)
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                    CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' 1111 1110
                    Call MsecDelay(0.8)
                 
                
                  
                
                
                LBA = LBA + 1
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                Call LabelMenu(0, rv0, 1)
                ClosePipe
               ' If rv0 = 1 Then
               '     TestResult = "PASS"
               ' End If
               
                If rv0 = 1 Then
                     rv1 = 1
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                     If LightOn <> 63 Then
                        rv1 = 2
                     End If
                 End If
               
                Tester.Print rv0, " \\ : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                Tester.Print rv1, " \\ : 1 GPO pass ,2 : GPO fail"
                   Call LabelMenu(2, rv1, rv0)
                   If rv1 = 2 Then
                   Tester.Label9.Caption = "GPO Fail"
                   End If
                
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
                XDWriteFail = XDWriteFail + 1
                TestResult = "XD_WF"
                ElseIf rv1 = READ_FAIL Then
                XDReadFail = XDReadFail + 1
                TestResult = "XD_RF"
                        
                    
                    
                ElseIf rv0 * rv1 = PASS Then
                     TestResult = "PASS"
                     
                Else
                      TestResult = "Bin2"
                End If
End Sub

Public Sub AU6395CLF21TestSub()
     
                    If PCI7248InitFinish = 0 Then
                      PCI7248Exist
                    End If
                
                   result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    
                    CardResult = DO_WritePort(card, Channel_P1A, &H80)  ' 1111 1110
                    Call MsecDelay(0.1)
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                    CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' 1111 1110
                    Call MsecDelay(1.3)
                 
                
                  
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                
                LBA = LBA + 1
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                Call LabelMenu(0, rv0, 1)
                ClosePipe
               ' If rv0 = 1 Then
               '     TestResult = "PASS"
               ' End If
                If rv0 = 1 Then
                     rv1 = 1
                    
                      
                     If LightOn <> 63 Then
                        rv1 = 2
                     End If
                 End If
               
               
                Tester.Print rv0, " \\ : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                Tester.Print rv1, " \\ : 1 GPO pass ,2 : GPO fail"
                   Call LabelMenu(2, rv1, rv0)
                   If rv1 = 2 Then
                   Tester.Label9.Caption = "GPO Fail"
                   End If
                
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
                XDWriteFail = XDWriteFail + 1
                TestResult = "XD_WF"
                ElseIf rv1 = READ_FAIL Then
                XDReadFail = XDReadFail + 1
                TestResult = "XD_RF"
                        
                    
                    
                ElseIf rv0 * rv1 = PASS Then
                     TestResult = "PASS"
                     
                Else
                      TestResult = "Bin2"
                End If
End Sub
Public Sub AU6395CLS10SortingSub()
     
                    If PCI7248InitFinish = 0 Then
                      PCI7248Exist
                    End If
                
                   result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    
                    CardResult = DO_WritePort(card, Channel_P1A, &H80)  ' 1111 1110
                    Call MsecDelay(0.1)
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                    CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' 1111 1110
                    
                Tester.Print "1st 2.4V R/W test"
                PowerSet (124)
                Call MsecDelay(2.3)
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                
                LBA = LBA + 1
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                Call LabelMenu(0, rv0, 1)
                ClosePipe
               ' If rv0 = 1 Then
               '     TestResult = "PASS"
               ' End If
                If rv0 = 1 Then
                     rv1 = 1
                    
                      
                     If LightOn <> 63 And LightOn <> 127 Then
                        rv1 = 2
                     End If
                 End If
               
               
                Tester.Print rv0, " \\ : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                Tester.Print rv1, " \\ : 1 GPO pass ,2 : GPO fail"
                   Call LabelMenu(2, rv1, rv0)
                   If rv1 = 2 Then
                   Tester.Label9.Caption = "GPO Fail"
                   End If
                
                Tester.Print "LBA="; LBA
            
                If rv0 = 0 Then
                  TestResult = "Bin2"
                  Exit Sub
                End If
                  
                    
                If rv0 * rv1 = PASS Then
                     TestResult = "PASS"
                     
                Else
                      TestResult = "Bin3"
                      Exit Sub
                End If
                
                '============================================2
                rv0 = 0
                rv1 = 1
                ReaderExist = 0
                Tester.Print "2nd 2.4V R/W test"
                  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    
                    CardResult = DO_WritePort(card, Channel_P1A, &H80)  ' 1111 1110
                    Call MsecDelay(0.1)
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                    CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' 1111 1110
                PowerSet (124)
                Call MsecDelay(2.3)
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                
                LBA = LBA + 1
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                Call LabelMenu(0, rv0, 1)
                ClosePipe
               ' If rv0 = 1 Then
               '     TestResult = "PASS"
               ' End If
                If rv0 = 1 Then
                     rv1 = 1
                    
                      
                     If LightOn <> 63 And LightOn <> 127 Then
                        rv1 = 2
                     End If
                 End If
               
               
                Tester.Print rv0, " \\ : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                Tester.Print rv1, " \\ : 1 GPO pass ,2 : GPO fail"
                   Call LabelMenu(2, rv1, rv0)
                   If rv1 = 2 Then
                   Tester.Label9.Caption = "GPO Fail"
                   End If
                
                Tester.Print "LBA="; LBA
            
         
                    
                If rv0 * rv1 = PASS Then
                     TestResult = "PASS"
                     
                Else
                      TestResult = "Bin4"
                      Exit Sub
                End If
                
                  '============================================3
                rv0 = 0
                rv1 = 1
                ReaderExist = 0
                Tester.Print "3rd 2.4V R/W test"
                 result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    
                    CardResult = DO_WritePort(card, Channel_P1A, &H80)  ' 1111 1110
                    Call MsecDelay(0.1)
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                    CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' 1111 1110
                PowerSet (124)
                Call MsecDelay(2.5)
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                
                LBA = LBA + 1
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                Call LabelMenu(0, rv0, 1)
                ClosePipe
               ' If rv0 = 1 Then
               '     TestResult = "PASS"
               ' End If
                If rv0 = 1 Then
                     rv1 = 1
                    
                      
                     If LightOn <> 63 And LightOn <> 127 Then
                        rv1 = 2
                     End If
                 End If
               
               
                Tester.Print rv0, " \\ : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                Tester.Print rv1, " \\ : 1 GPO pass ,2 : GPO fail"
                   Call LabelMenu(2, rv1, rv0)
                   If rv1 = 2 Then
                   Tester.Label9.Caption = "GPO Fail"
                   End If
                
                Tester.Print "LBA="; LBA
            
         
                    
                If rv0 * rv1 = PASS Then
                     TestResult = "PASS"
                     
                Else
                      TestResult = "Bin5"
                      Exit Sub
                End If
                
                
                
End Sub

Public Sub AU6395CLS11SortingSub()
     
                    If PCI7248InitFinish = 0 Then
                      PCI7248Exist
                    End If
                
                   result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    
                    CardResult = DO_WritePort(card, Channel_P1A, &H80)  ' 1111 1110
                    Call MsecDelay(0.1)
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                    CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' 1111 1110
                    
                Tester.Print "1st 2.37V R/W test"
                PowerSet (1237)
                Call MsecDelay(2.5)
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                
                LBA = LBA + 1
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                Call LabelMenu(0, rv0, 1)
                ClosePipe
               ' If rv0 = 1 Then
               '     TestResult = "PASS"
               ' End If
                If rv0 = 1 Then
                     rv1 = 1
                    
                      
                     If LightOn <> 63 And LightOn <> 127 Then
                        rv1 = 2
                     End If
                 End If
               
               
                Tester.Print rv0, " \\ : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                Tester.Print rv1, " \\ : 1 GPO pass ,2 : GPO fail"
                   Call LabelMenu(2, rv1, rv0)
                   If rv1 = 2 Then
                   Tester.Label9.Caption = "GPO Fail"
                   End If
                
                Tester.Print "LBA="; LBA
            
                If rv0 = 0 Then
                  TestResult = "Bin2"
                  Exit Sub
                End If
                  
                    
                If rv0 * rv1 = PASS Then
                     TestResult = "PASS"
                     
                Else
                      TestResult = "Bin3"
                      Exit Sub
                End If
                
                '============================================2
                rv0 = 0
                rv1 = 1
                ReaderExist = 0
                Tester.Print "2nd 2.37V R/W test"
                  result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    
                    CardResult = DO_WritePort(card, Channel_P1A, &H80)  ' 1111 1110
                    Call MsecDelay(0.1)
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                    CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' 1111 1110
                PowerSet (1237)
                Call MsecDelay(2.5)
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                
                LBA = LBA + 1
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                Call LabelMenu(0, rv0, 1)
                ClosePipe
               ' If rv0 = 1 Then
               '     TestResult = "PASS"
               ' End If
                If rv0 = 1 Then
                     rv1 = 1
                    
                      
                     If LightOn <> 63 And LightOn <> 127 Then
                        rv1 = 2
                     End If
                 End If
               
               
                Tester.Print rv0, " \\ : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                Tester.Print rv1, " \\ : 1 GPO pass ,2 : GPO fail"
                   Call LabelMenu(2, rv1, rv0)
                   If rv1 = 2 Then
                   Tester.Label9.Caption = "GPO Fail"
                   End If
                
                Tester.Print "LBA="; LBA
            
         
                    
                If rv0 * rv1 = PASS Then
                     TestResult = "PASS"
                     
                Else
                      TestResult = "Bin4"
                      Exit Sub
                End If
                
                  '============================================3
                rv0 = 0
                rv1 = 1
                ReaderExist = 0
                Tester.Print "3rd 2.37V R/W test"
                 result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    
                    CardResult = DO_WritePort(card, Channel_P1A, &H80)  ' 1111 1110
                    Call MsecDelay(0.1)
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                    CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' 1111 1110
                PowerSet (1237)
                Call MsecDelay(2.5)
                CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                
                LBA = LBA + 1
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                Call LabelMenu(0, rv0, 1)
                ClosePipe
               ' If rv0 = 1 Then
               '     TestResult = "PASS"
               ' End If
                If rv0 = 1 Then
                     rv1 = 1
                    
                      
                     If LightOn <> 63 And LightOn <> 127 Then
                        rv1 = 2
                     End If
                 End If
               
               
                Tester.Print rv0, " \\ : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                Tester.Print rv1, " \\ : 1 GPO pass ,2 : GPO fail"
                   Call LabelMenu(2, rv1, rv0)
                   If rv1 = 2 Then
                   Tester.Label9.Caption = "GPO Fail"
                   End If
                
                Tester.Print "LBA="; LBA
            
         
                    
                If rv0 * rv1 = PASS Then
                     TestResult = "PASS"
                     
                Else
                      TestResult = "Bin5"
                      Exit Sub
                End If
                
                
                
End Sub

 
Public Sub AU6980HLSTestSub()
                   
                   
   Tester.Print "=====5.5 V test"
                    Call PowerSet(3)
                    Call MsecDelay(0.1)
   
                    Call PowerSet(551)
                    Call MsecDelay(0.9)
                    
     If ChipName = "AU6980AN" Then
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
                   End If
                
                  
                
                
                LBA = LBA + 1
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                Call LabelMenu(0, rv0, 1)
                ClosePipe
               ' If rv0 = 1 Then
               '     TestResult = "PASS"
               ' End If
               
                If rv0 = 1 And ChipName = "AU6980AN" Then
                     rv1 = 1
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                     If LightOn <> 254 Then
                        rv1 = 2
                     End If
                 End If
               
                Tester.Print rv0, " \\ : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                If ChipName = "AU6980AN" Then
                   Tester.Print rv1, " \\ : 1 GPO pass ,2 : GPO fail"
                   Call LabelMenu(2, rv1, rv0)
                   If rv1 = 2 Then
                   Tester.Label9.Caption = "GPO Fail"
                   End If
                End If
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
                ElseIf rv0 = PASS Then
                     TestResult = "PASS"
                     If ChipName = "AU6980AN" Then
                        If rv0 * rv1 = 1 Then
                         TestResult = "PASS"
                        Else
                          TestResult = "Bin3"
                      End If
                    End If
                Else
                      TestResult = "Bin2"
              
                End If
                  Call PowerSet(3)
          If TestResult = "PASS" Then
             rv0 = 0
             rv1 = 0
             ReaderExist = 0
                 Tester.Print "=====4.5 V test"
               
                 Call MsecDelay(0.3)
                 Call PowerSet(451)
                    
                 Call MsecDelay(0.9)
     If ChipName = "AU6980AN" Then
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
                   End If
                
                  
                
                
                LBA = LBA + 1
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                Call LabelMenu(0, rv0, 1)
                ClosePipe
               ' If rv0 = 1 Then
               '     TestResult = "PASS"
               ' End If
               
                If rv0 = 1 And ChipName = "AU6980AN" Then
                     rv1 = 1
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                     If LightOn <> 254 Then
                        rv1 = 2
                     End If
                 End If
               
                Tester.Print rv0, " \\ : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                If ChipName = "AU6980AN" Then
                   Tester.Print rv1, " \\ : 1 GPO pass ,2 : GPO fail"
                   Call LabelMenu(2, rv1, rv0)
                   If rv1 = 2 Then
                   Tester.Label9.Caption = "GPO Fail"
                   End If
                End If
                Tester.Print "LBA="; LBA
            
            
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "XD_WF"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "MS_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "MS_RF"
                ElseIf rv0 = PASS Then
                     TestResult = "PASS"
                     If ChipName = "AU6980AN" Then
                        If rv0 * rv1 = 1 Then
                         TestResult = "PASS"
                        Else
                          TestResult = "Bin5"
                      End If
                    End If
                Else
                      TestResult = "Bin4"
              
                End If
                
          End If
                
End Sub

Public Sub AU6371DLS20SortingSub()
' 20081022 add light test for AU6980

 
                
                Call PowerSet2(0, "3.3", "0.5", 1, "1.77", "0.5", 1)
                Call MsecDelay(1.6)
            
                LBA = LBA + 1
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                 Tester.Print rv0, " \\ : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                ClosePipe
                If rv0 = 1 Then
                   ReaderExist = 0
                   rv0 = 0
                    Call PowerSet2(0, "3.3", "0.5", 1, "1.77", "0.5", 1)
                    Call MsecDelay(1.6)
            
                    LBA = LBA + 1
                    ClosePipe
                    rv0 = CBWTest_New(0, 1, "vid_058f")
                    ClosePipe
                     Tester.Print rv0, " \\ : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                End If
                
                 If rv0 = 1 Then
                   ReaderExist = 0
                   rv0 = 0
                    Call PowerSet2(0, "3.3", "0.5", 1, "1.77", "0.5", 1)
                    Call MsecDelay(1.6)
                    
                    LBA = LBA + 1
                    ClosePipe
                    rv0 = CBWTest_New(0, 1, "vid_058f")
                    ClosePipe
                     Tester.Print rv0, " \\ : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                End If
                
                 If rv0 = 1 Then
                   ReaderExist = 0
                   rv0 = 0
                    Call PowerSet2(0, "3.3", "0.5", 1, "1.77", "0.5", 1)
                    Call MsecDelay(1.6)
                    
                    LBA = LBA + 1
                    ClosePipe
                    rv0 = CBWTest_New(0, 1, "vid_058f")
                    ClosePipe
                     Tester.Print rv0, " \\ : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                End If
                
                Call LabelMenu(0, rv0, 1)
                ClosePipe
               
               
                  Call PowerSet2(2, "3.3", "0.5", 1, "1.77", "0.5", 1)
               
               
               
                
                Tester.Print "LBA="; LBA
            
            
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf UsbSpeedTestResult = 2 Then
                       
                    TestResult = "XD_WF"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv0 = PASS Then
                     TestResult = "PASS"
                      
                Else
                      TestResult = "Bin2"
                End If
End Sub
Public Sub AU6371DLS21SortingSub()
' 20081022 add light test for AU6980

        ' Call PowerSet2(2, "3.3", "0.5", 1, "1.77", "0.5", 1)
                 '  Call MsecDelay(0.4)
                    Call PowerSet2(0, "3.3", "0.5", 1, "1.77", "0.5", 1)
                    Call MsecDelay(2.3)
            
                LBA = LBA + 1
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                 Tester.Print rv0, " \\ : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                ClosePipe
                If rv0 = 1 Then
                   ReaderExist = 0
                   rv0 = 0
                    Call PowerSet2(2, "3.3", "0.5", 1, "1.77", "0.5", 1)
                   Call MsecDelay(0.4)
                    Call PowerSet2(1, "3.3", "0.5", 1, "1.77", "0.5", 1)
                    Call MsecDelay(1.8)
            
                    LBA = LBA + 1
                    ClosePipe
                    rv0 = CBWTest_New(0, 1, "vid_058f")
                    ClosePipe
                     Tester.Print rv0, " \\ : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                End If
                
                 If rv0 = 1 Then
                   ReaderExist = 0
                   rv0 = 0
                     Call PowerSet2(2, "3.3", "0.5", 1, "1.77", "0.5", 1)
                   Call MsecDelay(0.4)
                    Call PowerSet2(1, "3.3", "0.5", 1, "1.77", "0.5", 1)
                    Call MsecDelay(1.8)
                    LBA = LBA + 1
                    ClosePipe
                    rv0 = CBWTest_New(0, 1, "vid_058f")
                    ClosePipe
                     Tester.Print rv0, " \\ : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                End If
                
                 If rv0 = 1 Then
                   ReaderExist = 0
                   rv0 = 0
                   Call PowerSet2(2, "3.3", "0.5", 1, "1.77", "0.5", 1)
                   Call MsecDelay(0.4)
                    Call PowerSet2(1, "3.3", "0.5", 1, "1.77", "0.5", 1)
                    Call MsecDelay(1.8)
                    
                    LBA = LBA + 1
                    ClosePipe
                    rv0 = CBWTest_New(0, 1, "vid_058f")
                    ClosePipe
                     Tester.Print rv0, " \\ : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                End If
                
                Call LabelMenu(0, rv0, 1)
                ClosePipe
               
               
                  Call PowerSet2(2, "3.3", "0.5", 1, "1.77", "0.5", 1)
               
               
               
                
                Tester.Print "LBA="; LBA
            
            
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf UsbSpeedTestResult = 2 Then
                       
                    TestResult = "XD_WF"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv0 = PASS Then
                     TestResult = "PASS"
                      
                Else
                      TestResult = "Bin2"
                End If
End Sub
Public Sub AU6371DLS22SortingSub()
' 20081022 add light test for AU6980

        ' Call PowerSet2(2, "3.3", "0.5", 1, "1.77", "0.5", 1)
                 '  Call MsecDelay(0.4)
                    Call PowerSet2(0, "5.0", "0.5", 1, "1.77", "0.5", 1)
                    Call MsecDelay(2.3)
            
                LBA = LBA + 1
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                 Tester.Print rv0, " \\ : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                ClosePipe
                If rv0 = 1 Then
                   ReaderExist = 0
                   rv0 = 0
                    Call PowerSet2(2, "5.0", "0.5", 1, "1.77", "0.5", 1)
                   Call MsecDelay(0.4)
                    Call PowerSet2(1, "5.0", "0.5", 1, "1.77", "0.5", 1)
                    Call MsecDelay(1.8)
            
                    LBA = LBA + 1
                    ClosePipe
                    rv0 = CBWTest_New(0, 1, "vid_058f")
                    ClosePipe
                     Tester.Print rv0, " \\ : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                End If
                
                 If rv0 = 1 Then
                   ReaderExist = 0
                   rv0 = 0
                     Call PowerSet2(2, "5.0", "0.5", 1, "1.77", "0.5", 1)
                   Call MsecDelay(0.4)
                    Call PowerSet2(1, "5.0", "0.5", 1, "1.77", "0.5", 1)
                    Call MsecDelay(1.8)
                    LBA = LBA + 1
                    ClosePipe
                    rv0 = CBWTest_New(0, 1, "vid_058f")
                    ClosePipe
                     Tester.Print rv0, " \\ : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                End If
                
                 If rv0 = 1 Then
                   ReaderExist = 0
                   rv0 = 0
                   Call PowerSet2(2, "5.0", "0.5", 1, "1.77", "0.5", 1)
                   Call MsecDelay(0.4)
                    Call PowerSet2(1, "5.0", "0.5", 1, "1.77", "0.5", 1)
                    Call MsecDelay(1.8)
                    
                    LBA = LBA + 1
                    ClosePipe
                    rv0 = CBWTest_New(0, 1, "vid_058f")
                    ClosePipe
                     Tester.Print rv0, " \\ : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
                End If
                
                Call LabelMenu(0, rv0, 1)
                ClosePipe
               
               
                  Call PowerSet2(2, "5.0", "0.5", 1, "1.77", "0.5", 1)
               
               
               
                
                Tester.Print "LBA="; LBA
            
            
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf UsbSpeedTestResult = 2 Then
                       
                    TestResult = "XD_WF"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv0 = PASS Then
                     TestResult = "PASS"
                      
                Else
                      TestResult = "Bin2"
                End If
End Sub
Public Sub AU6371DLS40SortingSub()
' 20081022 add light test for AU6980

          Dim Sv1 As String
          Dim Sv2 As String
              '  Sv1 = "2.9"
               ' Sv2 = "1.40"
               
                  Sv1 = "3.3"
                 Sv2 = "1.4"
                 ' Call PowerSet2(2, "3.3", "0.5", 1, "1.77", "0.5", 1)
                 '  Call MsecDelay(0.4)
                    Call PowerSet2(0, Sv1, "0.5", 1, Sv2, "0.5", 1)
                    Call MsecDelay(1.8)
            
                LBA = LBA + 1
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                 Tester.Print rv0, " \\ : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                  Call LabelMenu(0, rv0, 1)
                 '===== Change Voltage to test
                 
                 Sv1 = "2.88"
                ClosePipe
                
                
                
                   ReaderExist = 0
                   
                    Call PowerSet2(2, Sv1, "0.5", 1, Sv2, "0.5", 1)
                   Call MsecDelay(0.4)
                    Call PowerSet2(1, Sv1, "0.5", 1, Sv2, "0.5", 1)
                    Call MsecDelay(1.6)
            
                    LBA = LBA + 1
                    ClosePipe
                    rv1 = CBWTest_New(0, rv0, "vid_058f")
                    ClosePipe
                    If rv1 = 0 Then
                       rv1 = 2
                    End If
                    
                     Tester.Print rv1, " \\ : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
               
                
                
                
                Call LabelMenu(1, rv1, rv0)
                ClosePipe
                If rv1 <> 1 Then
                 Tester.Label9.Caption = " MS sorting fail"
                End If
                  Call PowerSet2(2, "3.3", "0.5", 1, "1.77", "0.5", 1)
               
               
               
                
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
                    XDWriteFail = XDWriteFail + 1
                   TestResult = "XD_WF"
                ElseIf rv1 = READ_FAIL Then
                    XDReadFail = XDReadFail + 1
                   TestResult = "XD_WF"
                    
                ElseIf rv0 * rv1 = PASS Then
                     TestResult = "PASS"
                      
                Else
                      TestResult = "Bin2"
                End If
End Sub

Public Sub AU6371DLS41SortingSub()
' 20081022 add light test for AU6980

        
            
                LBA = LBA + 1
                ClosePipe
                rv0 = CBWTest_New_AU6371DLS41(0, 1, "vid_058f")
                 Tester.Print rv0, " \\ : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                  Call LabelMenu(0, rv0, 1)
                 '===== Change Voltage to test
                 
                
                ClosePipe
                
                
                
                 
                  Call PowerSet2(2, "3.3", "0.5", 1, "1.77", "0.5", 1)
               
               
               
                
                Tester.Print "LBA="; LBA
            
            
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf UsbSpeedTestResult = 2 Then
                       
                    TestResult = "XD_WF"
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
                     TestResult = "PASS"
                      
                Else
                      TestResult = "Bin2"
                End If
End Sub
Public Sub AU6980TestSub()
' 20081022 add light test for AU6980


     If ChipName = "AU6980AN" Then
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
        End If
                
         If ChipName = "AU6980" Then
                    If PCI7248InitFinish = 0 Then
                      PCI7248Exist
                    End If
                
                
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                    CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' 1111 1110
                     
        End If
                       
                
                
                  
               CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                
                LBA = LBA + 1
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                Call LabelMenu(0, rv0, 1)
                ClosePipe
               ' If rv0 = 1 Then
               '     TestResult = "PASS"
               ' End If
                If rv0 = 1 And ChipName = "AU6980" Then
                     rv1 = 1
                     
                     If LightOn <> 223 And LightOn <> 254 Then
                           
                            Call MsecDelay(0.8)
                            CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                             If LightOn <> 223 And LightOn <> 254 Then
                           
                             rv1 = 2
                            End If
                     End If
                 End If
              
               
                If rv0 = 1 And ChipName = "AU6980AN" Then
                     rv1 = 1
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                     If LightOn <> 254 Then
                        rv1 = 2
                     End If
                 End If
               
                Tester.Print rv0, " \\ : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                If ChipName = "AU6980AN" Or ChipName = "AU6980" Then
                   Tester.Print rv1, " \\ : 1 GPO pass ,2 : GPO fail"
                   Call LabelMenu(2, rv1, rv0)
                   If rv1 = 2 Then
                   Tester.Label9.Caption = "GPO Fail"
                   End If
                End If
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
                ElseIf rv0 = PASS Then
                     TestResult = "PASS"
                     If ChipName = "AU6980AN" Or ChipName = "AU6980" Then
                        If rv0 * rv1 = 1 Then
                         TestResult = "PASS"
                        Else
                          TestResult = "Bin3"
                      End If
                    End If
                Else
                      TestResult = "Bin2"
                End If
End Sub

Public Sub AU6980HLS20SortingSub()
' 20081022 add light test for AU6980

Dim ChipString As String
   
                Call MsecDelay(11)
   
                       
                
                rv0 = 0
                  
                ChipString = "vid_058f"
                 If GetDeviceName(ChipString) <> "" Then
                    rv0 = 1
                     
                  
                  End If
                  
                  
               Call LabelMenu(0, rv0, 1)
               
                If rv0 = UNKNOW Then
                   UnknowDeviceFail = UnknowDeviceFail + 1
                   TestResult = "UNKNOW"
                ElseIf rv0 = WRITE_FAIL Then
                    SDWriteFail = SDWriteFail + 1
                    TestResult = "SD_WF"
                ElseIf rv0 = READ_FAIL Then
                    SDReadFail = SDReadFail + 1
                    TestResult = "SD_RF"
                ElseIf rv0 = PASS Then
                     TestResult = "PASS"
                Else
                      TestResult = "Bin2"
                End If
End Sub

Public Sub AU6980OCP10TestSub()
' 20081022 add light test for AU6980


   Call PowerSet(1500)


     If ChipName = "AU6980AN" Then
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
        End If
                
         If ChipName = "AU6980OCP10" Then
                    If PCI7248InitFinish = 0 Then
                      PCI7248Exist
                    End If
                
                
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                    CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' 1111 1110
                     
        End If
                       
                
                
                  
               CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                
                LBA = LBA + 1
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                Call LabelMenu(0, rv0, 1)
                ClosePipe
               ' If rv0 = 1 Then
               '     TestResult = "PASS"
               ' End If
                If rv0 = 1 And ChipName = "AU6980OCP10" Then
                     rv1 = 1
                     
                     If LightOn <> 223 And LightOn <> 254 Then
                           
                            Call MsecDelay(0.8)
                            CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                             If LightOn <> 223 And LightOn <> 254 Then
                           
                             rv1 = 2
                            End If
                     End If
                 End If
              
               
                If rv0 = 1 And ChipName = "AU6980AN" Then
                     rv1 = 1
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                     If LightOn <> 254 Then
                        rv1 = 2
                     End If
                 End If
               
                Tester.Print rv0, " \\ : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                If ChipName = "AU6980AN" Or ChipName = "AU6980OCP10" Then
                   Tester.Print rv1, " \\ : 1 GPO pass ,2 : GPO fail"
                   Call LabelMenu(2, rv1, rv0)
                   If rv1 = 2 Then
                   Tester.Label9.Caption = "GPO Fail"
                   End If
                End If
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
                ElseIf rv0 = PASS Then
                     TestResult = "PASS"
                     If ChipName = "AU6980AN" Or ChipName = "AU6980OCP10" Then
                        If rv0 * rv1 = 1 Then
                         TestResult = "PASS"
                        Else
                          TestResult = "Bin3"
                      End If
                    End If
                Else
                      TestResult = "Bin2"
                End If
End Sub
Public Sub AU6985HLS10TestSub()
 
     If ChipName = "AU6985HLS10" Then
     
     
                    Call PowerSet(316)
                    If PCI7248InitFinish = 0 Then
                      PCI7248Exist
                    End If
                   Call MsecDelay(0.5)
                   result = DIO_PortConfig(card, Channel_P1B, OUTPUT_PORT)
            
                    CardResult = DO_WritePort(card, Channel_P1B, &H0)
                    
                    CardResult = DO_WritePort(card, Channel_P1A, &H1)  ' 1111 1110
                    Call MsecDelay(0.03)
                    result = DIO_PortConfig(card, Channel_P1B, INPUT_PORT)
                    CardResult = DO_WritePort(card, Channel_P1A, &H0)  ' 1111 1110
                    Call MsecDelay(0.7)
       End If
                
                  
                
                
                LBA = LBA + 1
                ClosePipe
                rv0 = CBWTest_New(0, 1, "vid_058f")
                Call LabelMenu(0, rv0, 1)
                ClosePipe
               ' If rv0 = 1 Then
               '     TestResult = "PASS"
               ' End If
               
                If rv0 = 1 And ChipName = "AU6980AN" Then
                     rv1 = 1
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                     If LightOn <> 254 Then
                        rv1 = 2
                     End If
                 End If
               
                Tester.Print rv0, " \\ : 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                If ChipName = "AU6980AN" Then
                   Tester.Print rv1, " \\ : 1 GPO pass ,2 : GPO fail"
                   Call LabelMenu(2, rv1, rv0)
                   If rv1 = 2 Then
                   Tester.Label9.Caption = "GPO Fail"
                   End If
                End If
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
                ElseIf rv0 = PASS Then
                     TestResult = "PASS"
                     If ChipName = "AU6980AN" Then
                        If rv0 * rv1 = 1 Then
                         TestResult = "PASS"
                        Else
                          TestResult = "Bin3"
                      End If
                    End If
                Else
                      TestResult = "Bin2"
                End If
End Sub

 






     
   
 


Public Sub AU6366ALF20TestSub()
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
                 
                 
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
                 Call MsecDelay(0.8)    'power on time
              
                '===============================================
                '  SD Card test
                '================================================
             
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                  
                  
                
            
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
                      
                    
                          
                         If rv0 <> 0 Then
                          If LightOn <> 31 Or LightOff <> 255 Then
                                    
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
               
                 If CardResult <> 0 Then
                    MsgBox "Set CF Card Detect Down Fail"
                    End
                 End If
                 
                 
               
                 ClosePipe
                 rv1 = CBWTest_New(0, rv0, ChipString)
                 ClosePipe
  
                  
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
                   
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
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
                 
                     CardResult = DO_WritePort(card, Channel_P1A, &H7B)
                 
               
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
               
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                
               
                If CardResult <> 0 Then
                    MsgBox "Set XD Card Detect Down Fail"
                    End
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
              
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
             
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect Down Fail"
                    End
                 End If
                 
                
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
Public Sub AU6371DLTest25()
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
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
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
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
                   
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
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
                
                
                      If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                        ElseIf rv0 = WRITE_FAIL Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
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

Public Sub AU6371DLTest26()
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
                      
                      
                       For i = 1 To 20
                      
                        If rv0 = 1 Then
                           
                            ClosePipe
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                               
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                             
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                             
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
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
                   
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
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

Public Sub AU6371DLTest27()
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
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on  for detect Normal mode
                 
                 Call MsecDelay(1.2)
                      
                      
                  If GetDeviceName(ChipString) = "" Then
                    rv0 = 0
                    GoTo AU6371DLResult
                  
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
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                               
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                             
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                             
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
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
                   
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
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

Public Sub AU6371DLS30SortingSub()

' add SD_Speed_Test at AU6371DLTest29
' add Ram_unstable Bin at Bin5 , the Bin are shrink to Bin2,Bin3,Bin4

     ' AU6371


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

              '  Call PowerSet2(1, "3.30", "0.50", 1, "3.30", "0.50", 1)
                
                Dim ChipString As String
                Dim i As Integer
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
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
                 Call MsecDelay(0.1)
                 
                 
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on  for detect Normal mode
                 
                 Call MsecDelay(1.5)
                      
                      
                  If GetDeviceName(ChipString) = "" Then
                    rv0 = 0
                    GoTo AU6371DLResult
                  
                  End If
                  
                      
                      
                      
                  '========= MS test ==============
                  
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect On Fail"
                    End
                 End If
                    Call MsecDelay(0.1)
                    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                
                
                 Call MsecDelay(0.03)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                
                  Call MsecDelay(AU6371EL_BootTime * 2)
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect Down Fail"
                    End
                 End If
                 If ChipName = "AU6371EL" Then
                   ReaderExist = 0
                 End If
                
                ClosePipe
                rv5 = CBWTest_New(0, 1, ChipString)
                
                ClosePipe
               Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                 
                '===============================================
                '  SD Card test
                '================================================
             '   If Left(ChipName, 10) = "AU6371DLF2" Then
                  
                
                 
                  
                  
                 If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect On Fail"
                    End
                 End If
                 
             '  End If
                   
                 '===========================================
                 'NO card test
                 '============================================
                     CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect On Fail"
                    End
                 End If
                    Call MsecDelay(0.5)
                 
  
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
                      
                      
                      rv0 = CBWTest_New(0, rv5, ChipString)
                      
                      If rv0 = 1 Then
                          rv0 = Read_SD_Speed_AU6371(0, 0, 18, "8Bits")
                          If rv0 <> 1 Then
                            rv0 = 2
                            Tester.Print "SD bus width Fail"
                          End If
                      End If
                      
                      ClosePipe
                      
                      
                      If rv0 = 1 Then
                      
                         For i = 1 To 20
                        
                          If rv0 = 1 Then
                             
                              ClosePipe
                               rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                                 
                               ClosePipe
                           End If
                   
                     
                          If rv0 = 1 Then
                            
                             ClosePipe
                              rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                               
                              ClosePipe
                          End If
                   
                
                          If rv0 = 1 Then
                            
                             ClosePipe
                              rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                                
                              ClosePipe
                          End If
                
                          If rv0 = 1 Then
                            
                             ClosePipe
                              rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                               
                              ClosePipe
                          End If
                           
                          If rv0 <> 1 Then
                          rv6 = 2  ' ram unstable
                          GoTo AU6371DLResult
                          End If
                             
                          Next
                          
                      End If
                        
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
                     
                          If rv0 <> 0 And rv5 = 1 Then
                           If LightOn <> &HBF Or LightOff <> &HFF Then
                                    
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
               
                
                 
              
                  rv1 = rv0 '----------- AU6371S3 dp not have CF slot
                 
               
                 
                 Call LabelMenu(1, rv1, rv0)
            
                   '   Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
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
                  
                  rv4 = rv3
                 '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
               
                Call LabelMenu(31, rv5, rv4)
                    
                
                '======== sorting
                
                  CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
AU6371DLResult:
                      If rv5 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                        ElseIf rv6 = 2 Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                            Tester.Label9.Caption = " ram unstable fail"
                      
                        ElseIf rv0 = WRITE_FAIL Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
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
                            TestResult = "CF_WF"
                        ElseIf rv2 = READ_FAIL Or rv3 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "CF_WF"
                         ElseIf rv4 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "CF_WF"
                            Tester.Label9.Caption = "XD Fail"
                        ElseIf rv4 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "CF_WF"
                           Tester.Label9.Caption = "XD Fail"
                       
                        ElseIf rv5 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                            Tester.Label9.Caption = "MS Fail"
                        ElseIf rv5 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_WF"
                           Tester.Label9.Caption = "MS Fail"
                        
                            
                            
                        ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub


Public Sub AU6371DLS32SortingSub()

' add SD_Speed_Test at AU6371DLTest29
' add Ram_unstable Bin at Bin5 , the Bin are shrink to Bin2,Bin3,Bin4

     ' AU6371


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

              '  Call PowerSet2(1, "3.30", "0.50", 1, "3.30", "0.50", 1)
                
                Dim ChipString As String
                Dim i As Integer
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
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
            '    Call PowerSet2(0, "3.13", "0.5", 1, "3.13", "0.5", 1)
                
                Call PowerSet2(0, "3.13", "0.065", 1, "3.13", "0.065", 1)
                
                 CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                    If CardResult <> 0 Then
                    MsgBox "Power off fail"
                    End
                 End If
                 Call MsecDelay(0.1)
                 
                 
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on  for detect Normal mode
                 
                 Call MsecDelay(2#)
                      
                      
                  If GetDeviceName(ChipString) = "" Then
                    rv0 = 0
                    GoTo AU6371DLResult
                  
                  End If
                  
                      
                      
                      
                  '========= MS test ==============
                  
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect On Fail"
                    End
                 End If
                    Call MsecDelay(0.1)
                    CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                
                
                 Call MsecDelay(0.03)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                
                  Call MsecDelay(AU6371EL_BootTime * 2)
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect Down Fail"
                    End
                 End If
                 If ChipName = "AU6371EL" Then
                   ReaderExist = 0
                 End If
                
                ClosePipe
                rv5 = CBWTest_New(0, 1, ChipString)
                
                ClosePipe
               Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                 
                '===============================================
                '  SD Card test
                '================================================
             '   If Left(ChipName, 10) = "AU6371DLF2" Then
                  
                
                 
                  
                  
                 If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect On Fail"
                    End
                 End If
                 
             '  End If
                   
                 '===========================================
                 'NO card test
                 '============================================
                   Call PowerSet2(1, "3.3", "0.5", 1, "3.3", "0.5", 1)
                   
                     CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                
                 If CardResult <> 0 Then
                    MsgBox "Set MSPro Card Detect On Fail"
                    End
                 End If
                    Call MsecDelay(0.5)
                 
  
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
                      
                      
                      rv0 = CBWTest_New(0, rv5, ChipString)
                      
                      If rv0 = 1 Then
                          rv0 = Read_SD_Speed_AU6371(0, 0, 18, "8Bits")
                          If rv0 <> 1 Then
                            rv0 = 2
                            Tester.Print "SD bus width Fail"
                          End If
                      End If
                      
                      ClosePipe
                      
                      
                      If rv0 = 1 Then
                      
                         For i = 1 To 20
                        
                          If rv0 = 1 Then
                             
                              ClosePipe
                               rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                                 
                               ClosePipe
                           End If
                   
                     
                          If rv0 = 1 Then
                            
                             ClosePipe
                              rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                               
                              ClosePipe
                          End If
                   
                
                          If rv0 = 1 Then
                            
                             ClosePipe
                              rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                                
                              ClosePipe
                          End If
                
                          If rv0 = 1 Then
                            
                             ClosePipe
                              rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                               
                              ClosePipe
                          End If
                           
                          If rv0 <> 1 Then
                          rv6 = 2  ' ram unstable
                          GoTo AU6371DLResult
                          End If
                             
                          Next
                          
                      End If
                        
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
                     
                          If rv0 <> 0 And rv5 = 1 Then
                           If LightOn <> &HBF Or LightOff <> &HFF Then
                                    
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
               
                
                 
              
                  rv1 = rv0 '----------- AU6371S3 dp not have CF slot
                 
               
                 
                 Call LabelMenu(1, rv1, rv0)
            
                   '   Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
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
                  
                  rv4 = rv3
                 '===============================================
                '  MS Pro Card test
                '================================================
              
                
                
               
                Call LabelMenu(31, rv5, rv4)
                    
                
                '======== sorting
                
                  CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                     Call PowerSet2(2, "3.3", "0.5", 1, "3.3", "0.5", 1)
                   
                
AU6371DLResult:
                      If rv5 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                        ElseIf rv6 = 2 Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                            Tester.Label9.Caption = " ram unstable fail"
                      
                        ElseIf rv0 = WRITE_FAIL Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
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
                            TestResult = "CF_WF"
                        ElseIf rv2 = READ_FAIL Or rv3 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "CF_WF"
                         ElseIf rv4 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "CF_WF"
                            Tester.Label9.Caption = "XD Fail"
                        ElseIf rv4 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "CF_WF"
                           Tester.Label9.Caption = "XD Fail"
                       
                        ElseIf rv5 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                            Tester.Label9.Caption = "MS Fail"
                        ElseIf rv5 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_WF"
                           Tester.Label9.Caption = "MS Fail"
                        
                            
                            
                        ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub
Public Sub AU6371DLS31SortingSub()

' add SD_Speed_Test at AU6371DLTest29
' add Ram_unstable Bin at Bin5 , the Bin are shrink to Bin2,Bin3,Bin4

     ' AU6371


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

            '    Call PowerSet2(1, "3.30", "0.50", 1, "3.30", "0.50", 1)
                
                Dim ChipString As String
                Dim i As Integer
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
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
                ' CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                '    If CardResult <> 0 Then
                '    MsgBox "Power off fail"
                '    End
                ' End If
                ' Call MsecDelay(0.05)
                 
                 
                 '    CardResult = DO_ReadPort(card, Channel_P1B, LightOFF)
                  
                '   If CardResult <> 0 Then
                '    MsgBox "Read light off fail"
                '    End
                '   End If
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on  for detect Normal mode
                 
                 Call MsecDelay(0.4)
                      
                      
                  If GetDeviceName(ChipString) = "" Then
                    rv0 = 0
                    GoTo AU6371DLResult
                  
                  End If
                  
                      
                      
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F - AU6371EL_SD)
                  
              '   Call MsecDelay(1.2 + AU6371EL_BootTime)  'power on time
              
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
                      
                      If rv0 = 1 Then
                          rv0 = Read_SD_Speed_AU6371(0, 0, 18, "8Bits")
                          If rv0 <> 1 Then
                            rv0 = 2
                            Tester.Print "SD bus width Fail"
                          End If
                      End If
                      
                      ClosePipe
                      
                      
                      If rv0 = 1 Then
                      
                         For i = 1 To 20
                        
                          If rv0 = 1 Then
                             
                              ClosePipe
                               rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                                 
                               ClosePipe
                           End If
                   
                     
                          If rv0 = 1 Then
                            
                             ClosePipe
                              rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                               
                              ClosePipe
                          End If
                   
                
                          If rv0 = 1 Then
                            
                             ClosePipe
                              rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                                
                              ClosePipe
                          End If
                
                          If rv0 = 1 Then
                            
                             ClosePipe
                              rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                               
                              ClosePipe
                          End If
                           
                          If rv0 <> 1 Then
                          rv6 = 2  ' ram unstable
                          GoTo AU6371DLResult
                          End If
                             
                          Next
                          
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
         
             
                  rv1 = rv0  '----------- NorthStar case: do not test CF
                 
               
                 
                 Call LabelMenu(1, rv1, rv0)
            
              '        Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
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
           
                rv2 = rv1
              
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
                  
                     rv4 = rv3
               
               
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
                 If ChipName = "AU6371EL" Then
                   ReaderExist = 0
                 End If
                
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                
                
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                
                '======== sorting
                rv7 = 1  'default to avoid error binning
                Dim OffTime As Single
                Dim StartTime As Single
                
                OffTime = 0.4
                StartTime = 1.6
                
               If rv5 = 1 Then  '-------- 1st
              
                
                   CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                   Call MsecDelay(OffTime)
                   CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                   Call MsecDelay(StartTime)
                
                    ReaderExist = 0
                    ClosePipe
                    rv7 = CBWTest_New(0, rv5, ChipString)
                    ClosePipe
                    Call LabelMenu(51, rv7, rv5)
                    Tester.Print rv7, "\\NS, MS sorting : 0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"";"
          
                End If
                
                If rv7 = 1 Then  '--------- 2nd
                   rv5 = rv7
                   
                     CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                   Call MsecDelay(OffTime)
                   CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                   Call MsecDelay(StartTime)
                
                    ReaderExist = 0
                    ClosePipe
                    rv7 = CBWTest_New(0, rv5, ChipString)
                    ClosePipe
                    Call LabelMenu(51, rv7, rv5)
                    Tester.Print rv7, "\\NS, MS sorting : 0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"";"
                
                End If
                
                  If rv7 = 1 Then  '--------- 3rd
                   rv5 = rv7
                   
                     CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                   Call MsecDelay(OffTime)
                   CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                   Call MsecDelay(StartTime)
                
                    ReaderExist = 0
                    ClosePipe
                    rv7 = CBWTest_New(0, rv5, ChipString)
                    ClosePipe
                    Call LabelMenu(51, rv7, rv5)
                    Tester.Print rv7, "\\NS, MS sorting : 0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"";"
                
                End If
                
                
                 If rv7 = 1 Then  '--------- 4th
                   rv5 = rv7
                   
                     CardResult = DO_WritePort(card, Channel_P1A, &HFF)
                   Call MsecDelay(OffTime)
                   CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                   Call MsecDelay(StartTime)
                
                    ReaderExist = 0
                    ClosePipe
                    rv7 = CBWTest_New(0, rv5, ChipString)
                    ClosePipe
                    Call LabelMenu(51, rv7, rv5)
                    Tester.Print rv7, "\\NS, MS sorting : 0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"";"
                
                End If
                
                
                
                
                  ' Call PowerSet2(0, "3.30", "0.50", 1, "3.30", "0.50", 1)
               '   CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
AU6371DLResult:
                      If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                        ElseIf rv6 = 2 Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                            Tester.Label9.Caption = " ram unstable fail"
                      
                        ElseIf rv0 = WRITE_FAIL Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
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
                            TestResult = "CF_WF"
                        ElseIf rv2 = READ_FAIL Or rv3 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "CF_WF"
                         ElseIf rv4 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "CF_WF"
                            Tester.Label9.Caption = "XD Fail"
                        ElseIf rv4 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "CF_WF"
                           Tester.Label9.Caption = "XD Fail"
                       
                        ElseIf rv5 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "CF_WF"
                            Tester.Label9.Caption = "MS Fail"
                        ElseIf rv5 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "CF_WF"
                           Tester.Label9.Caption = "MS Fail"
                        ElseIf rv7 <> 1 Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "XD_WF"
                            Tester.Label9.Caption = " 3.13V MS fail"
                            
                            
                        ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub

Public Sub AU6371DLF2ATestSub()

' add SD_Speed_Test at AU6371DLTest29
' add Ram_unstable Bin at Bin5 , the Bin are shrink to Bin2,Bin3,Bin4

     ' AU6371


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

                Call PowerSet2(1, "3.30", "0.50", 1, "3.30", "0.50", 1)
                
                Dim ChipString As String
                Dim i As Integer
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
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
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on  for detect Normal mode
                 
                 Call MsecDelay(1.2)
                      
                      
                  If GetDeviceName(ChipString) = "" Then
                    rv0 = 0
                    GoTo AU6371DLResult
                  
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
                      
                      If rv0 = 1 Then
                          rv0 = Read_SD_Speed_AU6371(0, 0, 18, "8Bits")
                          If rv0 <> 1 Then
                            rv0 = 2
                            Tester.Print "SD bus width Fail"
                          End If
                      End If
                      
                      ClosePipe
                      
                      
                      If rv0 = 1 Then
                      
                         For i = 1 To 20
                        
                          If rv0 = 1 Then
                             
                              ClosePipe
                               rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                                 
                               ClosePipe
                           End If
                   
                     
                          If rv0 = 1 Then
                            
                             ClosePipe
                              rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                               
                              ClosePipe
                          End If
                   
                
                          If rv0 = 1 Then
                            
                             ClosePipe
                              rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                                
                              ClosePipe
                          End If
                
                          If rv0 = 1 Then
                            
                             ClosePipe
                              rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                               
                              ClosePipe
                          End If
                           
                          If rv0 <> 1 Then
                          rv6 = 2  ' ram unstable
                          GoTo AU6371DLResult
                          End If
                             
                          Next
                          
                      End If
                        
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
                       If rv1 = 1 Then
                         
                          ClosePipe
                           rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                             Call LabelMenu(1, rv1, rv0)
                           ClosePipe
                       End If
                       
                         
                       If rv1 = 1 Then
                         
                          ClosePipe
                           rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                             Call LabelMenu(1, rv1, rv0)
                           ClosePipe
                       End If
                       
                    
                       If rv1 = 1 Then
                         
                          ClosePipe
                           rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                             Call LabelMenu(1, rv1, rv0)
                           ClosePipe
                       End If
                    
                       If rv1 = 1 Then
                         
                          ClosePipe
                           rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                             Call LabelMenu(1, rv1, rv0)
                           ClosePipe
                       End If
                      
                       If rv1 <> 1 Then ' Ram unstable
                      
                       rv6 = 2
                       GoTo AU6371DLResult
                       End If
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
                  
                     rv4 = rv3
               
               
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
                
                '======== sorting
                rv7 = 1
                
               If rv5 = 1 Then
                 Call PowerSet2(2, "3.30", "0.50", 1, "3.30", "0.50", 1)
                 Call MsecDelay(0.4)
                 Call PowerSet2(1, "3.13", "0.50", 1, "3.13", "0.50", 1)
                  Call MsecDelay(2.2)
                    ReaderExist = 0
                   rv7 = CBWTest_New(0, rv5, ChipString)
                    
                    Call LabelMenu(51, rv7, rv5)
                   Tester.Print rv7, "\\3.15V, MS : 0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"";"
                End If
                   Call PowerSet2(0, "3.30", "0.50", 1, "3.30", "0.50", 1)
                  CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
AU6371DLResult:
                      If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                        ElseIf rv6 = 2 Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                            Tester.Label9.Caption = " ram unstable fail"
                      
                        ElseIf rv0 = WRITE_FAIL Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
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
                            TestResult = "CF_WF"
                        ElseIf rv2 = READ_FAIL Or rv3 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "CF_WF"
                         ElseIf rv4 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "CF_WF"
                            Tester.Label9.Caption = "XD Fail"
                        ElseIf rv4 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "CF_WF"
                           Tester.Label9.Caption = "XD Fail"
                       
                        ElseIf rv5 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "CF_WF"
                            Tester.Label9.Caption = "MS Fail"
                        ElseIf rv5 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "CF_WF"
                           Tester.Label9.Caption = "MS Fail"
                        ElseIf rv7 <> 1 Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "XD_WF"
                            Tester.Label9.Caption = " 3.13V MS fail"
                            
                            
                        ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub



Public Sub AU6371DLTest29()

' add SD_Speed_Test at AU6371DLTest29
' add Ram_unstable Bin at Bin5 , the Bin are shrink to Bin2,Bin3,Bin4

     ' AU6371


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
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on  for detect Normal mode
                 
                 Call MsecDelay(1.2)
                      
                      
                  If GetDeviceName(ChipString) = "" Then
                    rv0 = 0
                    GoTo AU6371DLResult
                  
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
                      
                      If rv0 = 1 Then
                          rv0 = Read_SD_Speed_AU6371(0, 0, 18, "8Bits")
                          If rv0 <> 1 Then
                            rv0 = 2
                            Tester.Print "SD bus width Fail"
                          End If
                      End If
                      
                      ClosePipe
                      
                      
                      If rv0 = 1 Then
                      
                         For i = 1 To 20
                        
                          If rv0 = 1 Then
                             
                              ClosePipe
                               rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                                 
                               ClosePipe
                           End If
                   
                     
                          If rv0 = 1 Then
                            
                             ClosePipe
                              rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                               
                              ClosePipe
                          End If
                   
                
                          If rv0 = 1 Then
                            
                             ClosePipe
                              rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                                
                              ClosePipe
                          End If
                
                          If rv0 = 1 Then
                            
                             ClosePipe
                              rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                               
                              ClosePipe
                          End If
                           
                          If rv0 <> 1 Then
                          rv6 = 2  ' ram unstable
                          GoTo AU6371DLResult
                          End If
                             
                          Next
                          
                      End If
                        
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
                       If rv1 = 1 Then
                         
                          ClosePipe
                           rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                             Call LabelMenu(1, rv1, rv0)
                           ClosePipe
                       End If
                       
                         
                       If rv1 = 1 Then
                         
                          ClosePipe
                           rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                             Call LabelMenu(1, rv1, rv0)
                           ClosePipe
                       End If
                       
                    
                       If rv1 = 1 Then
                         
                          ClosePipe
                           rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                             Call LabelMenu(1, rv1, rv0)
                           ClosePipe
                       End If
                    
                       If rv1 = 1 Then
                         
                          ClosePipe
                           rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                             Call LabelMenu(1, rv1, rv0)
                           ClosePipe
                       End If
                      
                       If rv1 <> 1 Then ' Ram unstable
                      
                       rv6 = 2
                       GoTo AU6371DLResult
                       End If
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
                        ElseIf rv6 = 2 Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                            Tester.Label9.Caption = "ram unsatble Fail"
                        ElseIf rv0 = WRITE_FAIL Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
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
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                            Tester.Label9.Caption = "XD Fail"
                        ElseIf rv4 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                           Tester.Label9.Caption = "XD Fail"
                       
                        ElseIf rv5 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                            Tester.Label9.Caption = "MS Fail"
                        ElseIf rv5 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                           Tester.Label9.Caption = "MS Fail"
                       
                            
                        ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub

Public Sub AU6371DLF09TestSub()

' add SD_Speed_Test at AU6371DLTest29
' add Ram_unstable Bin at Bin5 , the Bin are shrink to Bin2,Bin3,Bin4

Dim ChipString As String
Dim i As Integer
Dim AU6371EL_SD As Byte
Dim AU6371EL_CF As Byte
Dim AU6371EL_XD As Byte
Dim AU6371EL_MS As Byte
Dim AU6371EL_MSP  As Byte
Dim AU6371EL_BootTime As Single

Dim HV_Done_Flag As Boolean
Dim HV_Result As String
Dim LV_Result As String

    If PCI7248InitFinish = 0 Then
          PCI7248Exist
    End If
    
    OldChipName = ""

    ChipString = "vid_058f"
    
    If PCI7248InitFinish = 0 Then
        PCI7248Exist
    End If
                   
Routine_Label:
 
    If Not HV_Done_Flag Then
        Call PowerSet2(0, "3.6", "0.5", 1, "3.6", "0.5", 1)
        Call MsecDelay(0.2)
        Tester.Print "AU6371DL : HV(3.6) Begin Test ..."
    Else
        Call PowerSet2(0, "3.3", "0.5", 1, "3.3", "0.5", 1)
        Call MsecDelay(0.2)
        Tester.Print vbCrLf & "AU6371DL : LV(3.3) Begin Test ..."
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
                 
    CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on  for detect Normal mode
    
    'Call MsecDelay(1.2)
    If WaitDevOn(ChipString) <> 1 Then
        rv0 = 0
        GoTo AU6371DLResult
    End If
                        
    '===============================================
    '  No Card test
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
    Call MsecDelay(0.04)
    
    CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
    If CardResult <> 0 Then
        MsgBox "Read light On fail"
        End
    End If
    
    ClosePipe
                      
    rv0 = CBWTest_New(0, 1, ChipString)
                      
    If rv0 = 1 Then
        rv0 = Read_SD_Speed_AU6371(0, 0, 18, "8Bits")
        If rv0 <> 1 Then
            rv0 = 2
            Tester.Print "SD bus width Fail"
        End If
    End If
                      
    ClosePipe
                      
                      
    If rv0 = 1 Then
    
        For i = 1 To 20
        
            If rv0 = 1 Then
                ClosePipe
                rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                ClosePipe
            End If
    
            If rv0 = 1 Then
                ClosePipe
                rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                ClosePipe
            End If
    
            If rv0 = 1 Then
                ClosePipe
                rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                ClosePipe
            End If
    
            If rv0 = 1 Then
                ClosePipe
                rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                ClosePipe
            End If
                           
            If rv0 <> 1 Then
                rv6 = 2  ' ram unstable
                GoTo AU6371DLResult
            End If
               
        Next
                          
    End If
                        
    If rv0 <> 0 Then
        If LightOn <> &HBF Or LightOff <> &HFF Then
            UsbSpeedTestResult = GPO_FAIL
            rv0 = 3
        End If
    End If
    
    Call LabelMenu(0, rv0, 1)   ' no card test fail
    
    Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    '===============================================
    '  CF Card test
    '================================================
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
    Call MsecDelay(0.1)

    ClosePipe
    rv1 = CBWTest_New(0, rv0, ChipString)
    ClosePipe

    If rv1 = 1 Then
        If rv1 = 1 Then
            ClosePipe
            rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
            Call LabelMenu(1, rv1, rv0)
            ClosePipe
        End If
        
        If rv1 = 1 Then
            ClosePipe
            rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
            Call LabelMenu(1, rv1, rv0)
            ClosePipe
        End If
        
        If rv1 = 1 Then
            ClosePipe
            rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
            Call LabelMenu(1, rv1, rv0)
            ClosePipe
        End If
        
        If rv1 = 1 Then
            ClosePipe
            rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
            Call LabelMenu(1, rv1, rv0)
            ClosePipe
        End If
        
        If rv1 <> 1 Then ' Ram unstable
            rv6 = 2
            GoTo AU6371DLResult
        End If
    End If
                      
                 
    Call LabelMenu(1, rv1, rv0)
            
    Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
    '===============================================
    '  SMC Card test  : stop these test for card not enough
    '================================================
    
    'Tester.print rv2, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    
    rv2 = rv1   ' to complete the SMC asbolish
               
              
    '===============================================
    '  XD Card test
    '================================================
    CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
    If CardResult <> 0 Then
        MsgBox "Set XD Card Detect On Fail"
        End
    End If
    
    Call MsecDelay(0.04)
    If rv2 = 1 Then
        CardResult = DO_WritePort(card, Channel_P1A, &H77)
    End If
    
    Call MsecDelay(0.1)
    
    If CardResult <> 0 Then
        MsgBox "Set XD Card Detect Down Fail"
        End
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
    
    Call MsecDelay(0.1)
    If CardResult <> 0 Then
       MsgBox "Set MSPro Card Detect Down Fail"
       End
    End If
    ClosePipe
    rv5 = CBWTest_New(0, rv4, ChipString)
    
    
    Call LabelMenu(31, rv5, rv4)
    Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
    ClosePipe
    
AU6371DLResult:

    CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
    Call PowerSet2(0, "0.0", "0.5", 1, "0.0", "0.5", 1)
    Call MsecDelay(0.2)
     
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
        GoTo Routine_Label
    Else
        If rv0 <> 1 Then
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

Public Sub AU6371DLS60SortingSub()

' add SD_Speed_Test at AU6371DLTest29
' add Ram_unstable Bin at Bin5 , the Bin are shrink to Bin2,Bin3,Bin4

     ' AU6371


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
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on  for detect Normal mode
                 
                 Call MsecDelay(1.2)
                      
                      
                  If GetDeviceName(ChipString) = "" Then
                    rv0 = 0
                    GoTo AU6371DLResult
                  
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
                      
               '       If rv0 = 1 Then
                '          rv0 = Read_SD_Speed_AU6371(0, 0, 18, "8Bits")
                '          If rv0 <> 1 Then
                 '           rv0 = 2
                  '          Tester.Print "SD bus width Fail"
                   '       End If
                   '   End If
                      
                      ClosePipe
                      
                      
                      If rv0 = 1 Then
                      
                         For i = 1 To 20
                        
                          If rv0 = 1 Then
                             
                              ClosePipe
                               rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                                 
                               ClosePipe
                           End If
                   
                     
                          If rv0 = 1 Then
                            
                             ClosePipe
                              rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                               
                              ClosePipe
                          End If
                   
                
                          If rv0 = 1 Then
                            
                             ClosePipe
                              rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                                
                              ClosePipe
                          End If
                
                          If rv0 = 1 Then
                            
                             ClosePipe
                              rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                               
                              ClosePipe
                          End If
                             
                          If rv0 <> 1 Then
                            If rv0 = 3 Then
                               rv6 = 2  ' ram unstable, data compare error
                            End If
                          GoTo AU6371DLResult
                          End If
                             
                          Next
                          
                      End If
                        
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
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                                    
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
                       If rv1 = 1 Then
                         
                          ClosePipe
                           rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                             Call LabelMenu(1, rv1, rv0)
                           ClosePipe
                       End If
                       
                         
                       If rv1 = 1 Then
                         
                          ClosePipe
                           rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                             Call LabelMenu(1, rv1, rv0)
                           ClosePipe
                       End If
                       
                    
                       If rv1 = 1 Then
                         
                          ClosePipe
                           rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                             Call LabelMenu(1, rv1, rv0)
                           ClosePipe
                       End If
                    
                       If rv1 = 1 Then
                         
                          ClosePipe
                           rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                             Call LabelMenu(1, rv1, rv0)
                           ClosePipe
                       End If
                      
                       If rv1 <> 1 Then ' Ram unstable
                      
                       rv6 = 2
                       GoTo AU6371DLResult
                       End If
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
                         ElseIf rv6 = 2 Then
                              TestResult = "PASS"
                             Tester.Label9.Caption = "ram unstable"
                        ElseIf rv0 = WRITE_FAIL Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
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
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                            Tester.Label9.Caption = "XD Fail"
                        ElseIf rv4 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                           Tester.Label9.Caption = "XD Fail"
                       
                        ElseIf rv5 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                            Tester.Label9.Caption = "MS Fail"
                        ElseIf rv5 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                           Tester.Label9.Caption = "MS Fail"
                       
                            
                        ElseIf rv6 = 2 Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub

Public Sub AU6371DLF2BTestSub()

' add SD_Speed_Test at AU6371DLTest29
' add Ram_unstable Bin at Bin5 , the Bin are shrink to Bin2,Bin3,Bin4

     ' AU6371


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
                
                Call PowerSet2(1, "3.3", "0.1", 1, "3.3", "0.1", 1)
                 
                Dim ChipString As String
                Dim i As Integer
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
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
               '  CardResult = DO_WritePort(card, Channel_P1A, &HFF)
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
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on  for detect Normal mode
                 
                 Call MsecDelay(1.2)
                      
                      
                  If GetDeviceName(ChipString) = "" Then
                    rv0 = 0
                    GoTo AU6371DLResult
                  
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
                      
                      If rv0 = 1 Then
                          rv0 = Read_SD_Speed_AU6371(0, 0, 18, "8Bits")
                          If rv0 <> 1 Then
                            rv0 = 2
                            Tester.Print "SD bus width Fail"
                          End If
                      End If
                      
                      ClosePipe
                      
                      
                      If rv0 = 1 Then
                      
                         For i = 1 To 20
                        
                          If rv0 = 1 Then
                             
                              ClosePipe
                               rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                                 
                               ClosePipe
                           End If
                   
                     
                          If rv0 = 1 Then
                            
                             ClosePipe
                              rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                               
                              ClosePipe
                          End If
                   
                
                          If rv0 = 1 Then
                            
                             ClosePipe
                              rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                                
                              ClosePipe
                          End If
                
                          If rv0 = 1 Then
                            
                             ClosePipe
                              rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                               
                              ClosePipe
                          End If
                           
                          If rv0 <> 1 Then
                          rv6 = 2  ' ram unstable
                          GoTo AU6371DLResult
                          End If
                             
                          Next
                          
                      End If
                        
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
                       If rv1 = 1 Then
                         
                          ClosePipe
                           rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                             Call LabelMenu(1, rv1, rv0)
                           ClosePipe
                       End If
                       
                         
                       If rv1 = 1 Then
                         
                          ClosePipe
                           rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                             Call LabelMenu(1, rv1, rv0)
                           ClosePipe
                       End If
                       
                    
                       If rv1 = 1 Then
                         
                          ClosePipe
                           rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                             Call LabelMenu(1, rv1, rv0)
                           ClosePipe
                       End If
                    
                       If rv1 = 1 Then
                         
                          ClosePipe
                           rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                             Call LabelMenu(1, rv1, rv0)
                           ClosePipe
                       End If
                      
                       If rv1 <> 1 Then ' Ram unstable
                      
                       rv6 = 2
                       GoTo AU6371DLResult
                       End If
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
                  
                     rv4 = rv3
               
               
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
               
                
               
                
                '  CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
AU6371DLResult:
                      If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                        ElseIf rv6 = 2 Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                            Tester.Label9.Caption = "ram unsatble Fail"
                        ElseIf rv0 = WRITE_FAIL Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
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
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                            Tester.Label9.Caption = "XD Fail"
                        ElseIf rv4 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                           Tester.Label9.Caption = "XD Fail"
                       
                        ElseIf rv5 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                            Tester.Label9.Caption = "MS Fail"
                        ElseIf rv5 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                           Tester.Label9.Caption = "MS Fail"
                       
                            
                        ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub
Public Sub AU6366C_IQCSub()
 
 Dim ChipString As String
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
                If Left(ChipName, 10) = "AU6371ELF2" Then
                  ChipName = "AU6371EL"
                  OldChipName = "AU6371ELF2"
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
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F - AU6371EL_SD)
                  
                 Call MsecDelay(1.2 + AU6371EL_BootTime)  'power on time
                 
                '===============================================
                '  SD Card test
                '================================================
                 
                
                  
                  
                 If CardResult <> 0 Then
                    MsgBox "Set SD Card Detect On Fail"
                    End
                 End If
                 
                   Call MsecDelay(0.01)
                 '===========================================
                 'NO card test
                 '============================================
                  
                  
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
                     Call MsecDelay(0.01)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
                            
                      If rv0 <> 0 Then
                        If LightOn <> &HBF Or LightOff <> &HFF Then
                                  
                        UsbSpeedTestResult = GPO_FAIL
                                
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
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
                   
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
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
                
              '       Print rv2, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
              
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
                   If Left(ChipName, 10) <> "AU6371GLF2" Then
                CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                 Call MsecDelay(0.03)
                 
                If CardResult <> 0 Then
                    MsgBox "Set MS Card Detect On Fail"
                    End
                End If
                
                If rv3 = 1 Then
                    CardResult = DO_WritePort(card, Channel_P1A, &H6F)
                End If
                Call MsecDelay(AU6371EL_BootTime)
                 If CardResult <> 0 Then
                    MsgBox "Set MS Card Detect Down Fail"
                    End
                 End If
                  If ChipName = "AU6371EL" Then
                   ReaderExist = 0
                 End If
                
                
                ClosePipe
                rv4 = CBWTest_New(0, rv3, ChipString)
                 ClosePipe
                Call LabelMenu(3, rv4, rv3)
                
                Else
                     rv4 = 1
                End If
               
                     Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
                If ChipName <> "AU6371EL" Then
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
                Else
                
                 rv5 = 1
                End If
                
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
Public Sub AU6371DLF2BTestSubOld()

' add SD_Speed_Test at AU6371DLTest29
' add Ram_unstable Bin at Bin5 , the Bin are shrink to Bin2,Bin3,Bin4

     ' AU6371


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
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on  for detect Normal mode
                 
                 Call MsecDelay(1.2)
                      
                      
                  If GetDeviceName(ChipString) = "" Then
                    rv0 = 0
                    GoTo AU6371DLResult
                  
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
                      
                      
                        If rv0 = 1 Then
                      
                        rv0 = CBWTest_New_AU6371Fail(0, 1, ChipString)
                      
                       End If
                      
                      If rv0 = 1 Then
                          rv0 = Read_SD_Speed_AU6371(0, 0, 18, "8Bits")
                          If rv0 <> 1 Then
                            rv0 = 2
                            Tester.Print "SD bus width Fail"
                          End If
                      End If
                      
                      ClosePipe
                      
                      
                      If rv0 = 1 Then
                      
                         For i = 1 To 20
                        
                          If rv0 = 1 Then
                             
                              ClosePipe
                               rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                                 
                               ClosePipe
                           End If
                   
                     
                          If rv0 = 1 Then
                            
                             ClosePipe
                              rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                               
                              ClosePipe
                          End If
                   
                
                          If rv0 = 1 Then
                            
                             ClosePipe
                              rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                                
                              ClosePipe
                          End If
                
                          If rv0 = 1 Then
                            
                             ClosePipe
                              rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                               
                              ClosePipe
                          End If
                           
                          If rv0 <> 1 Then
                          rv6 = 2  ' ram unstable
                          GoTo AU6371DLResult
                          End If
                             
                          Next
                          
                      End If
                        
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
                 
                 
                   If rv1 = 1 Then
                      
                        rv1 = CBWTest_New_AU6371Fail(0, 1, ChipString)
                      
                       End If
                 ClosePipe
  
                  
                  
                  
                 If rv1 = 1 Then
                       If rv1 = 1 Then
                         
                          ClosePipe
                           rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                             Call LabelMenu(1, rv1, rv0)
                           ClosePipe
                       End If
                       
                         
                       If rv1 = 1 Then
                         
                          ClosePipe
                           rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                             Call LabelMenu(1, rv1, rv0)
                           ClosePipe
                       End If
                       
                    
                       If rv1 = 1 Then
                         
                          ClosePipe
                           rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                             Call LabelMenu(1, rv1, rv0)
                           ClosePipe
                       End If
                    
                       If rv1 = 1 Then
                         
                          ClosePipe
                           rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                             Call LabelMenu(1, rv1, rv0)
                           ClosePipe
                       End If
                      
                       If rv1 <> 1 Then ' Ram unstable
                      
                       rv6 = 2
                       GoTo AU6371DLResult
                       End If
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
                  
                     rv4 = rv3
               
               
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
                
                ClosePipe
                
                   If rv5 = 1 Then
                      
                        rv5 = CBWTest_New_AU6371Fail(0, 1, ChipString)
                      
                       End If
                ClosePipe
                       
                         ClosePipe
                
                   If rv5 = 1 Then
                      
                        rv5 = CBWTest_New_AU6371Fail(0, 1, ChipString)
                      
                       End If
                ClosePipe
                
                  ClosePipe
                
                   If rv5 = 1 Then
                      
                        rv5 = CBWTest_New_AU6371Fail(0, 1, ChipString)
                      
                       End If
                ClosePipe
                
                  ClosePipe
                
                   If rv5 = 1 Then
                      
                        rv5 = CBWTest_New_AU6371Fail(0, 1, ChipString)
                      
                       End If
                ClosePipe
                       
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
               
                
               
                
                  CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
AU6371DLResult:
                      If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                        ElseIf rv6 = 2 Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                            Tester.Label9.Caption = "ram unsatble Fail"
                        ElseIf rv0 = WRITE_FAIL Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
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
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                            Tester.Label9.Caption = "XD Fail"
                        ElseIf rv4 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                           Tester.Label9.Caption = "XD Fail"
                       
                        ElseIf rv5 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                            Tester.Label9.Caption = "MS Fail"
                        ElseIf rv5 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                           Tester.Label9.Caption = "MS Fail"
                       
                            
                        ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub


Public Sub AU6371DLS50SortingSub()

' add SD_Speed_Test at AU6371DLTest29
' add Ram_unstable Bin at Bin5 , the Bin are shrink to Bin2,Bin3,Bin4

     ' AU6371


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
              
              Call PowerSet2(1, "3.3", "0.12", 1, "1.49", "0.12", 1)
              
             '  Call MsecDelay(0.8)
                Dim ChipString As String
                Dim i As Integer
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
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
              '   CardResult = DO_WritePort(card, Channel_P1A, &HFF)
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
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on  for detect Normal mode
                 
                 Call MsecDelay(1.2)
                      
                      
                  If GetDeviceName(ChipString) = "" Then
                    rv0 = 0
                    GoTo AU6371DLResult
                  
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
                      
                        If rv0 = 1 Then
                      ClosePipe
                        rv0 = CBWTest_New_AU6371Fail(0, 1, ChipString)
                      ClosePipe
                       End If
                       
                          If rv0 = 1 Then
                      ClosePipe
                        rv0 = CBWTest_New_AU6371Fail(0, 1, ChipString)
                      ClosePipe
                       End If
                       
                          If rv0 = 1 Then
                      ClosePipe
                        rv0 = CBWTest_New_AU6371Fail(0, 1, ChipString)
                      ClosePipe
                       End If
                       
                          If rv0 = 1 Then
                      ClosePipe
                        rv0 = CBWTest_New_AU6371Fail(0, 1, ChipString)
                      ClosePipe
                       End If
                       
                          If rv0 = 1 Then
                      ClosePipe
                        rv0 = CBWTest_New_AU6371Fail(0, 1, ChipString)
                      ClosePipe
                       End If
                      
                    '  If rv0 = 1 Then
                    '      rv0 = Read_SD_Speed_AU6371(0, 0, 18, "8Bits")
                    '      If rv0 <> 1 Then
                    '        rv0 = 2
                    '        Tester.Print "SD bus width Fail"
                    '      End If
                    '  End If
                      
                      
                      
                      
                      If rv0 = 1 Then
                      
                         For i = 1 To 20
                        
                          If rv0 = 1 Then
                             
                              ClosePipe
                               rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                                 
                               ClosePipe
                           End If
                   
                     
                          If rv0 = 1 Then
                            
                             ClosePipe
                              rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                               
                              ClosePipe
                          End If
                   
                
                          If rv0 = 1 Then
                            
                             ClosePipe
                              rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                                
                              ClosePipe
                          End If
                
                          If rv0 = 1 Then
                            
                             ClosePipe
                              rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                               
                              ClosePipe
                          End If
                           
                          If rv0 <> 1 Then
                          rv6 = 2  ' ram unstable
                          GoTo AU6371DLResult
                          End If
                             
                          Next
                          
                      End If
                   
                     
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
AU6371DLResult:
                      If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                        ElseIf rv6 = 2 Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                            Tester.Label9.Caption = "ram unsatble Fail"
                        ElseIf rv0 = WRITE_FAIL Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
                            Tester.Label9.Caption = "MS Fail"
                        ElseIf rv0 = READ_FAIL Then
                            SDReadFail = SDReadFail + 1
                            TestResult = "SD_RF"
                            Tester.Label9.Caption = "MS Fail"
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
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                            Tester.Label9.Caption = "XD Fail"
                        ElseIf rv4 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                           Tester.Label9.Caption = "XD Fail"
                       
                        ElseIf rv5 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                            Tester.Label9.Caption = "MS Fail"
                        ElseIf rv5 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                           Tester.Label9.Caption = "MS Fail"
                       
                            
                        ElseIf rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub
Public Sub AU6371DLTestOverCurrent()

' add SD_Speed_Test at AU6371DLTest29
' add Ram_unstable Bin at Bin5 , the Bin are shrink to Bin2,Bin3,Bin4

     ' AU6371


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
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on  for detect Normal mode
                 
                 Call MsecDelay(1.2)
                      
                      
                  If GetDeviceName(ChipString) = "" Then
                    rv0 = 0
                    GoTo AU6371DLResult
                  
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
                  
                   rv0 = CBWTest_New_no_card(0, 1, ChipString)
                    If rv0 = 1 Then
                       rv0 = SetOverCurrent(rv0)
                       If rv0 = 0 Then
                          rv0 = 2
                       End If
                     End If
                 ClosePipe
                    
                  
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
                      
                      
                      rv0 = CBWTest_New(0, rv0, ChipString)
                      
                      If rv0 = 1 Then
                          rv0 = Read_SD_Speed_AU6371(0, 0, 18, "8Bits")
                          If rv0 <> 1 Then
                            rv0 = 2
                            Tester.Print "SD bus width Fail"
                          End If
                      End If
                      
                      ClosePipe
                      
                      
                      If rv0 = 1 Then
                      
                         For i = 1 To 20
                        
                          If rv0 = 1 Then
                             
                              ClosePipe
                               rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                                 
                               ClosePipe
                           End If
                   
                     
                          If rv0 = 1 Then
                            
                             ClosePipe
                              rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                               
                              ClosePipe
                          End If
                   
                
                          If rv0 = 1 Then
                            
                             ClosePipe
                              rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                                
                              ClosePipe
                          End If
                
                          If rv0 = 1 Then
                            
                             ClosePipe
                              rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                               
                              ClosePipe
                          End If
                           
                          If rv0 <> 1 Then
                          rv6 = 2  ' ram unstable
                          GoTo AU6371DLResult
                          End If
                             
                          Next
                          
                      End If
                        
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
                          If LightOn <> &HBF Then
                                    
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
                       If rv1 = 1 Then
                         
                          ClosePipe
                           rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                             Call LabelMenu(1, rv1, rv0)
                           ClosePipe
                       End If
                       
                         
                       If rv1 = 1 Then
                         
                          ClosePipe
                           rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                             Call LabelMenu(1, rv1, rv0)
                           ClosePipe
                       End If
                       
                    
                       If rv1 = 1 Then
                         
                          ClosePipe
                           rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                             Call LabelMenu(1, rv1, rv0)
                           ClosePipe
                       End If
                    
                       If rv1 = 1 Then
                         
                          ClosePipe
                           rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                             Call LabelMenu(1, rv1, rv0)
                           ClosePipe
                       End If
                      
                       If rv1 <> 1 Then ' Ram unstable
                      
                       rv6 = 2
                       GoTo AU6371DLResult
                       End If
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
               
               
                   OpenPipe
                  If rv5 = 1 Then
                   
                    If rv5 = 1 Then
                     rv7 = SetOverCurrent(rv5)
                        If rv7 <> 1 Then
                          rv7 = 2
                        End If
                     End If
                     
                   If rv7 = 1 Then
                        rv7 = Read_OverCurrent(0, 0, 64)
                        If rv7 <> 1 Then
                        rv7 = 2
                        End If
                   End If
                   
                 End If
                    ClosePipe
                    Call LabelMenu(51, rv7, rv5)
                    If rv7 <> 1 Then
                    Tester.Label9.Caption = "Over Current fail"
                    Tester.Label2.Caption = "Over Current fail---"
                    End If
               
                   Tester.Print rv7, " \\OverCurrent :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
               
                  CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
AU6371DLResult:
                      If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                        ElseIf rv6 = 2 Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                            Tester.Label9.Caption = "ram unsatble Fail"
                        ElseIf rv0 = WRITE_FAIL Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
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
                            TestResult = "CF_WF"
                        ElseIf rv2 = READ_FAIL Or rv3 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "CF_WF"
                         ElseIf rv4 = WRITE_FAIL Or rv5 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "CF_WF"
                            Tester.Label9.Caption = "XD Fail"
                        ElseIf rv4 = READ_FAIL Or rv5 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "CF_WF"
                           Tester.Label9.Caption = "XD Fail"
                       
                        ElseIf rv7 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                            Tester.Label9.Caption = "MS Fail"
                        ElseIf rv7 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                           Tester.Label9.Caption = "MS Fail"
                       
                            
                        ElseIf rv7 * rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub
Public Sub AU6371DLTest28()
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
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on  for detect Normal mode
                 
                 Call MsecDelay(1.2)
                      
                      
                  If GetDeviceName(ChipString) = "" Then
                    rv0 = 0
                    GoTo AU6371DLResult
                  
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
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                               
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                             
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                             
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
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
                   
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
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
Public Sub AU6371DLSorting1()


ChipName = "AU6371DLF26"
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
               
               
               
               ' GPIB on
               
               
               
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
                Call PowerSet(3)
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
                 Call PowerSet(36)
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
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                               
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                             
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                             
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
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
                   
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
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
               
                
                Call PowerSet(3)
                
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
Public Sub AU6371DLTest()
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
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
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
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
                   
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
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
                   If Left(ChipName, 7) <> "AU6366C" And Left(ChipName, 10) <> "AU6371GLF2" And Left(OldChipName, 10) <> "AU6371HLF2" Then
                        CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                         Call MsecDelay(0.03)
                         
                        If CardResult <> 0 Then
                            MsgBox "Set MS Card Detect On Fail"
                            End
                        End If
                
                        If rv3 = 1 Then
                            CardResult = DO_WritePort(card, Channel_P1A, &H6F)
                        End If
                            Call MsecDelay(AU6371EL_BootTime)
                        If CardResult <> 0 Then
                           MsgBox "Set MS Card Detect Down Fail"
                           End
                        End If
                  If ChipName = "AU6371EL" Then
                   ReaderExist = 0
                 End If
                
                
                    ClosePipe
                    rv4 = CBWTest_New(0, rv3, ChipString)
                     ClosePipe
                    Call LabelMenu(3, rv4, rv3)
                
                Else
                     rv4 = 1
                End If
               
                     Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
                If ChipName <> "AU6371EL" Or OldChipName = "AU6371HLF2" Then
                
                
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
                Else
                
                rv5 = 1
                End If
                
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


Public Sub AU6371ELTest()
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
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
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
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
                   
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
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
                   If Left(ChipName, 7) <> "AU6366C" And Left(ChipName, 10) <> "AU6371GLF2" And Left(OldChipName, 10) <> "AU6371HLF2" Then
                        CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                         Call MsecDelay(0.03)
                         
                        If CardResult <> 0 Then
                            MsgBox "Set MS Card Detect On Fail"
                            End
                        End If
                
                        If rv3 = 1 Then
                            CardResult = DO_WritePort(card, Channel_P1A, &H6F)
                        End If
                            Call MsecDelay(AU6371EL_BootTime)
                        If CardResult <> 0 Then
                           MsgBox "Set MS Card Detect Down Fail"
                           End
                        End If
                  If ChipName = "AU6371EL" Then
                   ReaderExist = 0
                 End If
                
                
                    ClosePipe
                    rv4 = CBWTest_New(0, rv3, ChipString)
                     ClosePipe
                    Call LabelMenu(3, rv4, rv3)
                
                Else
                     rv4 = 1
                End If
               
                     Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
                If ChipName <> "AU6371EL" Or OldChipName = "AU6371HLF2" Then
                
                
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
                Else
                
                rv5 = 1
                End If
                
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





Public Sub AU6371ELS10Normal()
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
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
                If Left(ChipName, 10) = "AU6371ELS1" Then
                  ChipName = "AU6371EL"
                  OldChipName = "AU6371ELF2"
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
                
                
                      If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                        ElseIf rv0 = WRITE_FAIL Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
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

Public Sub AU6371ELTest25()
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
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
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
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
                   
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
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
                
                
                      If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                        ElseIf rv0 = WRITE_FAIL Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
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





Public Sub AU6371ELS20MSPro()
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
                
                 If Left(ChipName, 10) = "AU6371ELS2" Then
                  ChipName = "AU6371EL"
                  OldChipName = "AU6371ELF2"
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
                      
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
               
                
AU6371SortingResult:
                
                      If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                        ElseIf rv0 = WRITE_FAIL Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
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
Public Sub AU6371ELS11Ram()
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
                
                 If Left(ChipName, 10) = "AU6371ELS1" Then
                  ChipName = "AU6371EL"
                  OldChipName = "AU6371ELF2"
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
                 
                 
                    CardResult = DO_WritePort(card, Channel_P1A, &H7D)
                
                  
              Call MsecDelay(AU6371EL_BootTime * 2)
              
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
                     
                    rv0 = CBWTest_New(0, 1, ChipString)  ' for initial
                    
                    If rv0 = 1 Then
                    ClosePipe
                     rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       
                     ClosePipe
                
                  End If
                   
                 If rv0 = 1 Then
                   
                    ClosePipe
                     rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                      
                     ClosePipe
                 End If
                 
                   Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\123 pattern :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ' allen debug
              
              
              
                 If rv0 = 1 Then
                   
                    ClosePipe
                     rv2 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                     
                     ClosePipe
                 End If
              
                 If rv2 = 1 Then
                   
                    ClosePipe
                     rv2 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                        
                     ClosePipe
                 End If
                  
                           
                      
                      
                     ' rv0 = CBWTest_New(0, 1, ChipString)
                  
                     
                     Call LabelMenu(2, rv2, rv0)   ' no card test fail
                     
                 Tester.Print rv2, " \\012 pattern :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
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
                 If rv2 = 1 Then
                 
                     
                   CardResult = DO_WritePort(card, Channel_P1A, &H7F - AU6371EL_SD)
                  
                   Call MsecDelay(1.2 + AU6371EL_BootTime)  'power on time
              
                 
                  
                 End If
                 If CardResult <> 0 Then
                    MsgBox "Set CF Card Detect Down Fail"
                    End
                 End If
                 
                 If ChipName = "AU6371EL" Then
                   ReaderExist = 0
                 End If
                 
                 
                 
                  rv4 = CBWTest_New(0, 1, ChipString)  ' for initial
                  
                  If rv4 = 1 Then
                 
                      For i = 1 To 20
                      
                        
                           
                            ClosePipe
                             rv4 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                               
                             ClosePipe
                                                  
                 
                   
                        If rv4 = 1 Then
                          
                           ClosePipe
                            rv4 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
                 
              
                        If rv4 = 1 Then
                          
                           ClosePipe
                            rv4 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
              
                        If rv4 = 1 Then
                          
                           ClosePipe
                            rv4 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
              
                        If rv4 <> 1 Then
                        GoTo AU6371SortingResult
                        End If
                           
                        Next
                     End If
              
                 Call LabelMenu(3, rv4, rv2)
            
                      Tester.Print rv4, " \\ram unstable :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
AU6371SortingResult:
                
                      If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                        ElseIf rv0 = WRITE_FAIL Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
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
                       
                            
                        ElseIf rv2 * rv0 * rv4 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
                         
End Sub

Public Sub AU6371ELTest26()
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
                      
                      
                       For i = 1 To 20
                      
                        If rv0 = 1 Then
                           
                            ClosePipe
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                               
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                              
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
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
                   
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
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






Public Sub AU6371NLTest26()
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
                If Left(ChipName, 10) = "AU6371NLF2" Then
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
                      
                      
                       For i = 1 To 20
                      
                        If rv0 = 1 Then
                           
                            ClosePipe
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                              
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                               
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                              
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
                
                  rv1 = 1  '----------- AU6371S3 dp not have CF slot
                 
              
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

Public Sub AU6371SLTest26()
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
                If Left(ChipName, 10) = "AU6371SLF2" Then
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
                      
                      
                       For i = 1 To 20
                      
                        If rv0 = 1 Then
                           
                            ClosePipe
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                               
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                             
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                           
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                               
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
                
                  rv1 = 1  '----------- AU6371S3 dp not have CF slot
                 
              
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

Public Sub AU6371SLTest28()
Tester.Print "AU6371SL is NB mode"

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
                If Left(ChipName, 10) = "AU6371SLF2" Then
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
                 
                   
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on  for detect Normal mode
                 
                 Call MsecDelay(1.2)
                      
                      
                  If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    Tester.Print "NB mode Fail"
                    GoTo AU6371ELResult
                  
                  End If
                 
                 
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F - AU6371EL_SD)
                  
                 Call MsecDelay(1.2)     'power on time
              
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
                      
                      
                       For i = 1 To 20
                      
                        If rv0 = 1 Then
                           
                            ClosePipe
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                               
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                             
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                           
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                               
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
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                                    
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
                
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
              
                 Call LabelMenu(1, rv1, rv0)
            
                '      Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
              '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
                   CardResult = DO_WritePort(card, Channel_P1A, &H7A) '1010 SD + SMC
               
                    Call MsecDelay(0.2)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H7B)
                    Call MsecDelay(0.2)
               
                  If CardResult <> 0 Then
                     MsgBox "Set SMC Card Detect Down Fail"
                     End
                  End If
                 
                   OpenPipe
                    rv2 = ReInitial(0)
                    ClosePipe
                 
                 ClosePipe
                 rv2 = CBWTest_New(0, rv1, "vid_058f")
                 Call LabelMenu(21, rv2, rv1)
                
                     Tester.Print rv2, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
              
               
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H73) '0011 XD + SMC
               
                    Call MsecDelay(0.2)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                    Call MsecDelay(0.2)
               
                  If CardResult <> 0 Then
                     MsgBox "Set XD Card Detect Down Fail"
                     End
                  End If
                 
                   OpenPipe
                    rv3 = ReInitial(0)
                    ClosePipe
               
                  
                 
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
               
                '===============================================
                '  MS Pro Card test
                '================================================
                   rv4 = rv3  ' for MS
                
                
                   CardResult = DO_WritePort(card, Channel_P1A, &H57) ' XD + MSPro
               
                    Call MsecDelay(0.2)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                    Call MsecDelay(0.2)
               
                  If CardResult <> 0 Then
                     MsgBox "Set MSpro Card Detect Down Fail"
                     End
                  End If
                 
                   OpenPipe
                    rv5 = ReInitial(0)
                    ClosePipe
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

Public Sub AU6371SLTest29()
Tester.Print "AU6371SL is NB mode"

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
                If Left(ChipName, 10) = "AU6371SLF2" Then
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
                 
                   
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on  for detect Normal mode
                 
                 Call MsecDelay(1.2)
                      
                      
                  If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    Tester.Print "NB mode Fail"
                    GoTo AU6371ELResult
                  
                  End If
                 
                 
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F - AU6371EL_SD)
                  
                 Call MsecDelay(1.2)     'power on time
              
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
                         
                       If rv0 = 1 Then
                          rv0 = Read_SD_Speed_AU6371(0, 0, 18, "8Bits")
                          If rv0 <> 1 Then
                            rv0 = 2
                            Tester.Print "SD bus width Fail"
                          End If
                      End If
                      
                      
                      ClosePipe
                      
                        If rv0 = 1 Then
                       For i = 1 To 20
                      
                        If rv0 = 1 Then
                           
                            ClosePipe
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                               
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
                       
                        If rv0 <> 1 Then
                        rv6 = 2
                        GoTo AU6371ELResult
                        End If
                           
                        Next
                        
                     End If
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
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                                    
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
                
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
              
                 Call LabelMenu(1, rv1, rv0)
            
                '      Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
              '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
                   CardResult = DO_WritePort(card, Channel_P1A, &H7A) '1010 SD + SMC
               
                    Call MsecDelay(0.2)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H7B)
                    Call MsecDelay(0.2)
               
                  If CardResult <> 0 Then
                     MsgBox "Set SMC Card Detect Down Fail"
                     End
                  End If
                 
                   OpenPipe
                    rv2 = ReInitial(0)
                    ClosePipe
                 
                 ClosePipe
                 rv2 = CBWTest_New(0, rv1, "vid_058f")
                 Call LabelMenu(21, rv2, rv1)
                
                     Tester.Print rv2, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
              
               
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H73) '0011 XD + SMC
               
                    Call MsecDelay(0.2)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                    Call MsecDelay(0.2)
               
                  If CardResult <> 0 Then
                     MsgBox "Set XD Card Detect Down Fail"
                     End
                  End If
                 
                   OpenPipe
                    rv3 = ReInitial(0)
                    ClosePipe
               
                  
                 
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
               
                '===============================================
                '  MS Pro Card test
                '================================================
                   rv4 = rv3  ' for MS
                
                
                   CardResult = DO_WritePort(card, Channel_P1A, &H57) ' XD + MSPro
               
                    Call MsecDelay(0.2)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                    Call MsecDelay(0.2)
               
                  If CardResult <> 0 Then
                     MsgBox "Set MSpro Card Detect Down Fail"
                     End
                  End If
                 
                   OpenPipe
                    rv5 = ReInitial(0)
                    ClosePipe
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
                        ElseIf rv6 = 2 Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                            Tester.Label9.Caption = "ram unsatble Fail"
                        ElseIf rv0 = WRITE_FAIL Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
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
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                            Tester.Label9.Caption = "XD Fail"
                        ElseIf rv4 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                           Tester.Label9.Caption = "XD Fail"
                       
                        ElseIf rv5 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                            Tester.Label9.Caption = "MS Fail"
                        ElseIf rv5 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                           Tester.Label9.Caption = "MS Fail"
                       
                            
                        ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub
Public Sub AU6371SLTest27()
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
                If Left(ChipName, 10) = "AU6371SLF2" Then
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
                 
                   
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on  for detect Normal mode
                 
                 Call MsecDelay(1.2)
                      
                      
                  If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                 
                 
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
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                               
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                             
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                           
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                               
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
                
                  rv1 = 1  '----------- AU6371S3 dp not have CF slot
                 
              
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

Public Sub AU6371TLTest26()
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
                If Left(ChipName, 10) = "AU6371TLF2" Then
                  ChipName = "AU6371EL"
                  OldChipName = "AU6371ELF2"
                End If
                 
                 If Left(ChipName, 10) = "AU6371HLF2" Then
                  ChipName = "AU6371EL"
                  OldChipName = "AU6371HLF2"
                End If
                
                
                  ChipString = "vid_18e3"
               
                
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
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                              
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                               
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                               
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                               
                            ClosePipe
                        End If
              
                        If rv0 <> 1 Then
                        GoTo AU6371TLResult
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
                
                  rv1 = 1  '----------- AU6371S3 dp not have CF slot
                 
              
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
                  ClosePipe
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
                
AU6371TLResult:
                      If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                        ElseIf rv0 = WRITE_FAIL Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
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

Public Sub AU6371TLTest28()
Tester.Print "AU6371TL is NB mode"
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
                If Left(ChipName, 10) = "AU6371TLF2" Then
                  ChipName = "AU6371EL"
                  OldChipName = "AU6371ELF2"
                End If
                 
                 If Left(ChipName, 10) = "AU6371HLF2" Then
                  ChipName = "AU6371EL"
                  OldChipName = "AU6371HLF2"
                End If
                
                
                  ChipString = "vid_18e3"
               
                
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
                 
                 
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on  for detect Normal mode
                 
                 Call MsecDelay(1.2)
                      
                      
                  If GetDeviceName(ChipString) <> "" Then
                  
                    Tester.Print "NB mode Fail"
                    rv0 = 0
                    GoTo AU6371TLResult
                  
                  End If
                 
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
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
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      ClosePipe
                      
                      
                       For i = 1 To 20
                      
                        If rv0 = 1 Then
                           
                            ClosePipe
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                              
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                               
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                               
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                               
                            ClosePipe
                        End If
              
                        If rv0 <> 1 Then
                        GoTo AU6371TLResult
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
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                                    
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                     End If
                     
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  CF Card test
                '================================================
                
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
              
                 Call LabelMenu(1, rv1, rv0)
            
                '      Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
              '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
                   CardResult = DO_WritePort(card, Channel_P1A, &H7A) '1010 SD + SMC
               
                    Call MsecDelay(0.2)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H7B)
                    Call MsecDelay(0.2)
               
                  If CardResult <> 0 Then
                     MsgBox "Set SMC Card Detect Down Fail"
                     End
                  End If
                 
                   OpenPipe
                    rv2 = ReInitial(0)
                    ClosePipe
                 
                 ClosePipe
                 rv2 = CBWTest_New(0, rv1, "vid_058f")
                 Call LabelMenu(21, rv2, rv1)
                
                     Tester.Print rv2, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
              
               
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H73) '0011 XD + SMC
               
                    Call MsecDelay(0.2)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                    Call MsecDelay(0.2)
               
                  If CardResult <> 0 Then
                     MsgBox "Set XD Card Detect Down Fail"
                     End
                  End If
                 
                   OpenPipe
                    rv3 = ReInitial(0)
                    ClosePipe
               
                  
                 
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
               
                '===============================================
                '  MS Pro Card test
                '================================================
                   rv4 = rv3  ' for MS
                
                
                   CardResult = DO_WritePort(card, Channel_P1A, &H57) ' XD + MSPro
               
                    Call MsecDelay(0.2)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                    Call MsecDelay(0.2)
               
                  If CardResult <> 0 Then
                     MsgBox "Set MSpro Card Detect Down Fail"
                     End
                  End If
                 
                   OpenPipe
                    rv5 = ReInitial(0)
                    ClosePipe
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                
                
                  CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
AU6371TLResult:
                      If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                        ElseIf rv0 = WRITE_FAIL Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
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

Public Sub AU6371TLTest29()
Tester.Print "AU6371TL is NB mode"
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
                If Left(ChipName, 10) = "AU6371TLF2" Then
                  ChipName = "AU6371EL"
                  OldChipName = "AU6371ELF2"
                End If
                 
                 If Left(ChipName, 10) = "AU6371HLF2" Then
                  ChipName = "AU6371EL"
                  OldChipName = "AU6371HLF2"
                End If
                
                
                  ChipString = "vid_18e3"
               
                
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
                 
                 
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on  for detect Normal mode
                 
                 Call MsecDelay(1.2)
                      
                      
                  If GetDeviceName(ChipString) <> "" Then
                  
                    Tester.Print "NB mode Fail"
                    rv0 = 0
                    GoTo AU6371TLResult
                  
                  End If
                 
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
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
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      
                      
                       If rv0 = 1 Then
                          rv0 = Read_SD_Speed_AU6371(0, 0, 18, "8Bits")
                          If rv0 <> 1 Then
                            rv0 = 2
                            Tester.Print "SD bus width Fail"
                          End If
                      End If
                      
                      ClosePipe
                      
                      
                       If rv0 = 1 Then
                       For i = 1 To 20
                      
                        If rv0 = 1 Then
                           
                            ClosePipe
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                               
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
                       
                        If rv0 <> 1 Then
                        rv6 = 2
                        GoTo AU6371TLResult
                        End If
                           
                        Next
                        
                     End If
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
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                                    
                          UsbSpeedTestResult = GPO_FAIL
                          rv0 = 3
                          End If
                        End If
                     End If
                     
                     
                     Call LabelMenu(0, rv0, 1)   ' no card test fail
                     
                 Tester.Print rv0, " \\SD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  CF Card test
                '================================================
                
                  rv1 = rv0  '----------- AU6371S3 dp not have CF slot
                 
              
                 Call LabelMenu(1, rv1, rv0)
            
                '      Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
              '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
                   CardResult = DO_WritePort(card, Channel_P1A, &H7A) '1010 SD + SMC
               
                    Call MsecDelay(0.2)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H7B)
                    Call MsecDelay(0.2)
               
                  If CardResult <> 0 Then
                     MsgBox "Set SMC Card Detect Down Fail"
                     End
                  End If
                 
                   OpenPipe
                    rv2 = ReInitial(0)
                    ClosePipe
                 
                 ClosePipe
                 rv2 = CBWTest_New(0, rv1, "vid_058f")
                 Call LabelMenu(21, rv2, rv1)
                
                     Tester.Print rv2, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
              
               
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H73) '0011 XD + SMC
               
                    Call MsecDelay(0.2)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                    Call MsecDelay(0.2)
               
                  If CardResult <> 0 Then
                     MsgBox "Set XD Card Detect Down Fail"
                     End
                  End If
                 
                   OpenPipe
                    rv3 = ReInitial(0)
                    ClosePipe
               
                  
                 
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
               
                '===============================================
                '  MS Pro Card test
                '================================================
                   rv4 = rv3  ' for MS
                
                
                   CardResult = DO_WritePort(card, Channel_P1A, &H57) ' XD + MSPro
               
                    Call MsecDelay(0.2)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                    Call MsecDelay(0.2)
               
                  If CardResult <> 0 Then
                     MsgBox "Set MSpro Card Detect Down Fail"
                     End
                  End If
                 
                   OpenPipe
                    rv5 = ReInitial(0)
                    ClosePipe
                ClosePipe
                rv5 = CBWTest_New(0, rv4, ChipString)
                Call LabelMenu(31, rv5, rv4)
                     Tester.Print rv5, " \\MSpro :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                ClosePipe
                
                
                  CardResult = DO_WritePort(card, Channel_P1A, &HFF)   ' Close power
                
AU6371TLResult:
                      If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                        ElseIf rv6 = 2 Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                            Tester.Label9.Caption = "ram unsatble Fail"
                        ElseIf rv0 = WRITE_FAIL Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
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
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                            Tester.Label9.Caption = "XD Fail"
                        ElseIf rv4 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                           Tester.Label9.Caption = "XD Fail"
                       
                        ElseIf rv5 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                            Tester.Label9.Caption = "MS Fail"
                        ElseIf rv5 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                           Tester.Label9.Caption = "MS Fail"
                       
                            
                        ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub
 

Public Sub AU6371TLTest27()
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
                If Left(ChipName, 10) = "AU6371TLF2" Then
                  ChipName = "AU6371EL"
                  OldChipName = "AU6371ELF2"
                End If
                 
                 If Left(ChipName, 10) = "AU6371HLF2" Then
                  ChipName = "AU6371EL"
                  OldChipName = "AU6371HLF2"
                End If
                
                
                  ChipString = "vid_18e3"
               
                
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
                 
                 
                  CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on  for detect Normal mode
                 
                 Call MsecDelay(1.2)
                      
                      
                  If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371TLResult
                  
                  End If
                 
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
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                              
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                               
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                               
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                               
                            ClosePipe
                        End If
              
                        If rv0 <> 1 Then
                        GoTo AU6371TLResult
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
                
                  rv1 = 1  '----------- AU6371S3 dp not have CF slot
                 
              
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
                  ClosePipe
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
                
AU6371TLResult:
                      If rv0 = UNKNOW Then
                           UnknowDeviceFail = UnknowDeviceFail + 1
                           TestResult = "UNKNOW"
                        ElseIf rv0 = WRITE_FAIL Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
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



Public Sub AU6371HLTest28()
Tester.Print "AU6371HL is NB mode"
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
                 
                 
               
                 
                     
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on  for detect Normal mode
                 
                 Call MsecDelay(1.2)
                      
                      
                  If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    Tester.Print "NB mode Fail"
                    GoTo AU6371ELResult
                  
                  End If
                  
                  
                 
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F - AU6371EL_SD)
                  
                
                     Call MsecDelay(1.2)
                     
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
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                               
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                              
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
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                                    
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
                    CardResult = DO_WritePort(card, Channel_P1A, &H7C)  'SD+CF
                    Call MsecDelay(0.2)
                    
                    CardResult = DO_WritePort(card, Channel_P1A, &H7D)  'SMC
                    Call MsecDelay(0.2)
                    OpenPipe
                    rv1 = ReInitial(0)
                    ClosePipe
                  
              
               
                    ClosePipe
                      rv1 = CBWTest_New(0, rv0, ChipString)
                    ClosePipe
  
                  
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
                   
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
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
             '      CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
             '    If CardResult <> 0 Then
             '        MsgBox "Set SMC Card Detect On Fail"
             '        End
              '    End If
                  
              '    Call MsecDelay(0.01)
              '   If rv1 = 1 Then
              '       CardResult = DO_WritePort(card, Channel_P1A, &H7B)
              '   End If
               
              '  If CardResult <> 0 Then
              '       MsgBox "Set SMC Card Detect Down Fail"
              '       End
              '    End If
                 
               '  Call MsecDelay(1.2 + AU6371EL_BootTime)  'power on time
               '    If ChipName = "AU6371EL" Then
               '    ReaderExist = 0
               '  End If
                 
               '  ClosePipe
               '  rv2 = CBWTest_New(0, rv1, "vid_058f")
                ' Call LabelMenu(21, rv2, rv1)
                
                 '    Tester.Print rv2, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
              
               rv2 = rv1
                 '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H75) '0011 XD + CF
               
                    Call MsecDelay(0.2)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                    Call MsecDelay(0.2)
               
                  If CardResult <> 0 Then
                     MsgBox "Set XD Card Detect Down Fail"
                     End
                  End If
                 
                   OpenPipe
                    rv3 = ReInitial(0)
                    ClosePipe
               
                  
                 
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
               
                '===============================================
                '  MS Pro Card test
                '================================================
                   rv4 = rv3  ' for MS
                
                
                   CardResult = DO_WritePort(card, Channel_P1A, &H57) ' XD + MSPro
               
                    Call MsecDelay(0.2)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                    Call MsecDelay(0.2)
               
                  If CardResult <> 0 Then
                     MsgBox "Set MSpro Card Detect Down Fail"
                     End
                  End If
                 
                   OpenPipe
                    rv5 = ReInitial(0)
                    ClosePipe
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

Public Sub AU6371HLTest29()
Tester.Print "AU6371HL is NB mode"
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
                 
                 
               
                 
                     
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on  for detect Normal mode
                 
                 Call MsecDelay(1.2)
                      
                      
                  If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    Tester.Print "NB mode Fail"
                    GoTo AU6371ELResult
                  
                  End If
                  
                  
                 
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F - AU6371EL_SD)
                  
                
                     Call MsecDelay(1.2)
                     
                     CardResult = DO_ReadPort(card, Channel_P1B, LightOn)
                      If CardResult <> 0 Then
                          MsgBox "Read light On fail"
                          End
                      End If
                     
                     
                             
                           
                      ClosePipe
                      
                      
                      rv0 = CBWTest_New(0, 1, ChipString)
                      
                      If rv0 = 1 Then
                          rv0 = Read_SD_Speed_AU6371(0, 0, 18, "8Bits")
                          If rv0 <> 1 Then
                            rv0 = 2
                            Tester.Print "SD bus width Fail"
                          End If
                      End If
                      ClosePipe
                      
                      
                          If rv0 = 1 Then
                       For i = 1 To 20
                      
                        If rv0 = 1 Then
                           
                            ClosePipe
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                               
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
                       
                        If rv0 <> 1 Then
                        rv6 = 2
                        GoTo AU6371ELResult
                        End If
                           
                        Next
                        
                     End If
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
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                                    
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
                    CardResult = DO_WritePort(card, Channel_P1A, &H7C)  'SD+CF
                    Call MsecDelay(0.2)
                    
                    CardResult = DO_WritePort(card, Channel_P1A, &H7D)  'SMC
                    Call MsecDelay(0.2)
                    OpenPipe
                    rv1 = ReInitial(0)
                    ClosePipe
                  
              
               
                    ClosePipe
                      rv1 = CBWTest_New(0, rv0, ChipString)
                    ClosePipe
  
                  
                If rv1 = 1 Then
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
                   
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                   
                  If rv1 <> 1 Then
                      rv6 = 2
                      GoTo AU6371ELResult
                  End If
                
                
                End If
                
                
                 
               Else
                  rv1 = 1  '----------- AU6371S3 dp not have CF slot
                 
               End If
                 
                 Call LabelMenu(1, rv1, rv0)
            
                      Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
               '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
             '      CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
             '    If CardResult <> 0 Then
             '        MsgBox "Set SMC Card Detect On Fail"
             '        End
              '    End If
                  
              '    Call MsecDelay(0.01)
              '   If rv1 = 1 Then
              '       CardResult = DO_WritePort(card, Channel_P1A, &H7B)
              '   End If
               
              '  If CardResult <> 0 Then
              '       MsgBox "Set SMC Card Detect Down Fail"
              '       End
              '    End If
                 
               '  Call MsecDelay(1.2 + AU6371EL_BootTime)  'power on time
               '    If ChipName = "AU6371EL" Then
               '    ReaderExist = 0
               '  End If
                 
               '  ClosePipe
               '  rv2 = CBWTest_New(0, rv1, "vid_058f")
                ' Call LabelMenu(21, rv2, rv1)
                
                 '    Tester.Print rv2, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
              
               rv2 = rv1
                 '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H75) '0011 XD + CF
               
                    Call MsecDelay(0.2)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                    Call MsecDelay(0.2)
               
                  If CardResult <> 0 Then
                     MsgBox "Set XD Card Detect Down Fail"
                     End
                  End If
                 
                   OpenPipe
                    rv3 = ReInitial(0)
                    ClosePipe
               
                  
                 
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
               
                '===============================================
                '  MS Pro Card test
                '================================================
                   rv4 = rv3  ' for MS
                
                
                   CardResult = DO_WritePort(card, Channel_P1A, &H57) ' XD + MSPro
               
                    Call MsecDelay(0.2)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                    Call MsecDelay(0.2)
               
                  If CardResult <> 0 Then
                     MsgBox "Set MSpro Card Detect Down Fail"
                     End
                  End If
                 
                   OpenPipe
                    rv5 = ReInitial(0)
                    ClosePipe
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
                        ElseIf rv6 = 2 Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                            Tester.Label9.Caption = "ram unsatble Fail"
                        ElseIf rv0 = WRITE_FAIL Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
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
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                            Tester.Label9.Caption = "XD Fail"
                        ElseIf rv4 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                           Tester.Label9.Caption = "XD Fail"
                       
                        ElseIf rv5 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                            Tester.Label9.Caption = "MS Fail"
                        ElseIf rv5 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                           Tester.Label9.Caption = "MS Fail"
                       
                            
                        ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub

Public Sub AU6371HLTest27()
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
                 
                 
               
                 
                     
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on  for detect Normal mode
                 
                 Call MsecDelay(1.2)
                      
                      
                  If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                  
                  
                 
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
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                               
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                              
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
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
                   
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
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
             '      CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
             '    If CardResult <> 0 Then
             '        MsgBox "Set SMC Card Detect On Fail"
             '        End
              '    End If
                  
              '    Call MsecDelay(0.01)
              '   If rv1 = 1 Then
              '       CardResult = DO_WritePort(card, Channel_P1A, &H7B)
              '   End If
               
              '  If CardResult <> 0 Then
              '       MsgBox "Set SMC Card Detect Down Fail"
              '       End
              '    End If
                 
               '  Call MsecDelay(1.2 + AU6371EL_BootTime)  'power on time
               '    If ChipName = "AU6371EL" Then
               '    ReaderExist = 0
               '  End If
                 
               '  ClosePipe
               '  rv2 = CBWTest_New(0, rv1, "vid_058f")
                ' Call LabelMenu(21, rv2, rv1)
                
                 '    Tester.Print rv2, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
              
                If rv1 = 1 Then
                rv2 = 1
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
                If rv1 = 1 Then
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

Public Sub AU6371GLTest27()
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
                 
                 
               
                 
                     
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on  for detect Normal mode
                 
                 Call MsecDelay(1.2)
                      
                      
                  If GetDeviceName(ChipString) = "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                  
                  
                 
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
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                               
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                              
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
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
                   
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
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
             '      CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
             '    If CardResult <> 0 Then
             '        MsgBox "Set SMC Card Detect On Fail"
             '        End
              '    End If
                  
              '    Call MsecDelay(0.01)
              '   If rv1 = 1 Then
              '       CardResult = DO_WritePort(card, Channel_P1A, &H7B)
              '   End If
               
              '  If CardResult <> 0 Then
              '       MsgBox "Set SMC Card Detect Down Fail"
              '       End
              '    End If
                 
               '  Call MsecDelay(1.2 + AU6371EL_BootTime)  'power on time
               '    If ChipName = "AU6371EL" Then
               '    ReaderExist = 0
               '  End If
                 
               '  ClosePipe
               '  rv2 = CBWTest_New(0, rv1, "vid_058f")
                ' Call LabelMenu(21, rv2, rv1)
                
                 '    Tester.Print rv2, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
              
                If rv1 = 1 Then
                rv2 = 1
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
                If rv1 = 1 Then
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

Public Sub AU6371GLTest29()
' add SD_Speed_Test at AU6371GLTest29
' add Ram_unstable Bin at Bin5 , the Bin are shrink to Bin2,Bin3,Bin4
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
                 
                 
               
                 
                     
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on  for detect Normal mode
                 
                 Call MsecDelay(1.2)
                      
                      
                  If GetDeviceName(ChipString) = "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                  
                  
                 
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
                      
                       If rv0 = 1 Then
                          rv0 = Read_SD_Speed_AU6371(0, 0, 18, "8Bits")
                          If rv0 <> 1 Then
                            rv0 = 2
                            Tester.Print "SD bus width Fail"
                          End If
                      End If
                      ClosePipe
                      
                      If rv0 = 1 Then
                       For i = 1 To 20
                      
                        If rv0 = 1 Then
                           
                            ClosePipe
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                               
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
                       
                        If rv0 <> 1 Then
                        rv6 = 2
                        GoTo AU6371ELResult
                        End If
                           
                        Next
                        
                     End If
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
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
                   
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                   
                  If rv1 <> 1 Then
                      rv6 = 2
                      GoTo AU6371ELResult
                  End If
                
                
                End If
                
                 
               Else
                  rv1 = 1  '----------- AU6371S3 dp not have CF slot
                 
               End If
                 
                 Call LabelMenu(1, rv1, rv0)
            
                      Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
             '      CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
             '    If CardResult <> 0 Then
             '        MsgBox "Set SMC Card Detect On Fail"
             '        End
              '    End If
                  
              '    Call MsecDelay(0.01)
              '   If rv1 = 1 Then
              '       CardResult = DO_WritePort(card, Channel_P1A, &H7B)
              '   End If
               
              '  If CardResult <> 0 Then
              '       MsgBox "Set SMC Card Detect Down Fail"
              '       End
              '    End If
                 
               '  Call MsecDelay(1.2 + AU6371EL_BootTime)  'power on time
               '    If ChipName = "AU6371EL" Then
               '    ReaderExist = 0
               '  End If
                 
               '  ClosePipe
               '  rv2 = CBWTest_New(0, rv1, "vid_058f")
                ' Call LabelMenu(21, rv2, rv1)
                
                 '    Tester.Print rv2, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
              
                If rv1 = 1 Then
                rv2 = 1
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
                If rv1 = 1 Then
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
                        ElseIf rv6 = 2 Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                            Tester.Label9.Caption = "ram unsatble Fail"
                        ElseIf rv0 = WRITE_FAIL Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
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
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                            Tester.Label9.Caption = "XD Fail"
                        ElseIf rv4 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                           Tester.Label9.Caption = "XD Fail"
                       
                        ElseIf rv5 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                            Tester.Label9.Caption = "MS Fail"
                        ElseIf rv5 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                           Tester.Label9.Caption = "MS Fail"
                       
                            
                        ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub

Public Sub AU6371GLTest28()
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
                 
                 
               
                 
                     
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on  for detect Normal mode
                 
                 Call MsecDelay(1.2)
                      
                      
                  If GetDeviceName(ChipString) = "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                  
                  
                 
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
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                               
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                              
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
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
                   
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
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
             '      CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                  
             '    If CardResult <> 0 Then
             '        MsgBox "Set SMC Card Detect On Fail"
             '        End
              '    End If
                  
              '    Call MsecDelay(0.01)
              '   If rv1 = 1 Then
              '       CardResult = DO_WritePort(card, Channel_P1A, &H7B)
              '   End If
               
              '  If CardResult <> 0 Then
              '       MsgBox "Set SMC Card Detect Down Fail"
              '       End
              '    End If
                 
               '  Call MsecDelay(1.2 + AU6371EL_BootTime)  'power on time
               '    If ChipName = "AU6371EL" Then
               '    ReaderExist = 0
               '  End If
                 
               '  ClosePipe
               '  rv2 = CBWTest_New(0, rv1, "vid_058f")
                ' Call LabelMenu(21, rv2, rv1)
                
                 '    Tester.Print rv2, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
              
                If rv1 = 1 Then
                rv2 = 1
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
                If rv1 = 1 Then
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

Public Sub AU6371ELTest28()
Tester.Print "AU6371EL is NB mode"
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
                 
                 
               
                 
                     
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on  for detect Normal mode
                 
                 Call MsecDelay(1.2)   ' For NB power on
                      
                      
                  If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    Tester.Print "NB Mode Fail"
                    GoTo AU6371ELResult
                  
                  End If
                 
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F - AU6371EL_SD)
                  
                 Call MsecDelay(1.5)     'pull down SD card to power on time
              
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
                      
                      
                       For i = 1 To 20
                      
                        If rv0 = 1 Then
                           
                            ClosePipe
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                              
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
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                                    
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
                    CardResult = DO_WritePort(card, Channel_P1A, &H7C)  'SD+CF
                    Call MsecDelay(0.2)
                    
                    CardResult = DO_WritePort(card, Channel_P1A, &H7D)  'SMC
                    Call MsecDelay(0.2)
                    OpenPipe
                    rv1 = ReInitial(0)
                    ClosePipe
                  
              
               
                    ClosePipe
                      rv1 = CBWTest_New(0, rv0, ChipString)
                    ClosePipe
  
                  
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
                   
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
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
                   CardResult = DO_WritePort(card, Channel_P1A, &H75) '1001 CF + SMC
               
                    Call MsecDelay(0.2)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H7B)
                    Call MsecDelay(0.2)
               
                  If CardResult <> 0 Then
                     MsgBox "Set SMC Card Detect Down Fail"
                     End
                  End If
                 
                   OpenPipe
                    rv2 = ReInitial(0)
                    ClosePipe
                 
                 ClosePipe
                 rv2 = CBWTest_New(0, rv1, "vid_058f")
                 Call LabelMenu(21, rv2, rv1)
                
                     Tester.Print rv2, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
              
               
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H73) '0011 XD + SMC
               
                    Call MsecDelay(0.2)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                    Call MsecDelay(0.2)
               
                  If CardResult <> 0 Then
                     MsgBox "Set XD Card Detect Down Fail"
                     End
                  End If
                 
                   OpenPipe
                    rv3 = ReInitial(0)
                    ClosePipe
               
                  
                 
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
               
                '===============================================
                '  MS Pro Card test
                '================================================
                   rv4 = rv3  ' for MS
                
                
                   CardResult = DO_WritePort(card, Channel_P1A, &H57) ' XD + MSPro
               
                    Call MsecDelay(0.2)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                    Call MsecDelay(0.2)
               
                  If CardResult <> 0 Then
                     MsgBox "Set MSpro Card Detect Down Fail"
                     End
                  End If
                 
                   OpenPipe
                    rv5 = ReInitial(0)
                    ClosePipe
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

Public Sub AU6371ELTest29()
Tester.Print "AU6371EL is NB mode"
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
                 
                 
               
                 
                     
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on  for detect Normal mode
                 
                 Call MsecDelay(1.2)   ' For NB power on
                      
                      
                  If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    Tester.Print "NB Mode Fail"
                    GoTo AU6371ELResult
                  
                  End If
                 
                  CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                   If CardResult <> 0 Then
                    MsgBox "Read light off fail"
                    End
                   End If
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F - AU6371EL_SD)
                  
                 Call MsecDelay(1.5)     'pull down SD card to power on time
              
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
                      If rv0 = 1 Then
                         rv0 = Read_SD_Speed_AU6371(0, 0, 18, "8Bits")
                         If rv0 <> 1 Then
                            rv0 = 2
                            Tester.Print "SD bus width Fail"
                         End If
                         
                      End If
                      ClosePipe
                      
                      
                      If rv0 = 1 Then
                      
                       For i = 1 To 20
                      
                        If rv0 = 1 Then
                           
                            ClosePipe
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
              
                        If rv0 <> 1 Then
                        rv6 = 2
                        GoTo AU6371ELResult
                        End If
                           
                        Next
                        
                        
                        End If
                        
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
                          If LightOn <> &HBF Or LightOff <> &HFF Then
                                    
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
                    CardResult = DO_WritePort(card, Channel_P1A, &H7C)  'SD+CF
                    Call MsecDelay(0.2)
                    
                    CardResult = DO_WritePort(card, Channel_P1A, &H7D)  'SMC
                    Call MsecDelay(0.2)
                    OpenPipe
                    rv1 = ReInitial(0)
                    ClosePipe
                  
              
               
                    ClosePipe
                      rv1 = CBWTest_New(0, rv0, ChipString)
                    ClosePipe
  
                 If rv1 = 1 Then
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
                   
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
              
                    If rv1 <> 1 Then
                        rv6 = 2
                        GoTo AU6371ELResult
                    End If
                           
              
                
                 End If
                
                 
               Else
                  rv1 = 1  '----------- AU6371S3 dp not have CF slot
                 
               End If
                 
                 Call LabelMenu(1, rv1, rv0)
            
                      Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
                   CardResult = DO_WritePort(card, Channel_P1A, &H75) '1001 CF + SMC
               
                    Call MsecDelay(0.2)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H7B)
                    Call MsecDelay(0.2)
               
                  If CardResult <> 0 Then
                     MsgBox "Set SMC Card Detect Down Fail"
                     End
                  End If
                 
                   OpenPipe
                    rv2 = ReInitial(0)
                    ClosePipe
                 
                 ClosePipe
                 rv2 = CBWTest_New(0, rv1, "vid_058f")
                 Call LabelMenu(21, rv2, rv1)
                
                     Tester.Print rv2, " \\SMC :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
              
               
               
              
                '===============================================
                '  XD Card test
                '================================================
                  CardResult = DO_WritePort(card, Channel_P1A, &H73) '0011 XD + SMC
               
                    Call MsecDelay(0.2)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H77)
                    Call MsecDelay(0.2)
               
                  If CardResult <> 0 Then
                     MsgBox "Set XD Card Detect Down Fail"
                     End
                  End If
                 
                   OpenPipe
                    rv3 = ReInitial(0)
                    ClosePipe
               
                  
                 
                rv3 = CBWTest_New(0, rv2, ChipString)
                 ClosePipe
                Call LabelMenu(2, rv3, rv2)
                 
                     Tester.Print rv3, " \\XD :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                
               
                '===============================================
                '  MS Pro Card test
                '================================================
                   rv4 = rv3  ' for MS
                
                
                   CardResult = DO_WritePort(card, Channel_P1A, &H57) ' XD + MSPro
               
                    Call MsecDelay(0.2)
                
                    CardResult = DO_WritePort(card, Channel_P1A, &H5F)
                    Call MsecDelay(0.2)
               
                  If CardResult <> 0 Then
                     MsgBox "Set MSpro Card Detect Down Fail"
                     End
                  End If
                 
                   OpenPipe
                    rv5 = ReInitial(0)
                    ClosePipe
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
                        ElseIf rv6 = 2 Then
                            MSWriteFail = MSWriteFail + 1
                            TestResult = "MS_WF"
                            Tester.Label9.Caption = "ram unsatble Fail"
                        ElseIf rv0 = WRITE_FAIL Then
                            SDWriteFail = SDWriteFail + 1
                            TestResult = "SD_WF"
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
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                            Tester.Label9.Caption = "XD Fail"
                        ElseIf rv4 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                           Tester.Label9.Caption = "XD Fail"
                       
                        ElseIf rv5 = WRITE_FAIL Then
                            XDWriteFail = XDWriteFail + 1
                            TestResult = "XD_WF"
                            Tester.Label9.Caption = "MS Fail"
                        ElseIf rv5 = READ_FAIL Then
                            XDReadFail = XDReadFail + 1
                            TestResult = "XD_RF"
                           Tester.Label9.Caption = "MS Fail"
                       
                            
                        ElseIf rv5 * rv4 * rv3 * rv2 * rv1 * rv0 = PASS Then
                             TestResult = "PASS"
                        Else
                            TestResult = "Bin2"
                          
                        End If
End Sub

Public Sub AU6371ELTest27()
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
                 
                 
               
                 
                     
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on  for detect Normal mode
                 
                 Call MsecDelay(1.2)
                      
                      
                  If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                  
                  
                 
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
                             rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                               
                             ClosePipe
                         End If
                 
                   
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
                 
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                              
                            ClosePipe
                        End If
              
                        If rv0 = 1 Then
                          
                           ClosePipe
                            rv0 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                              
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
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
                   
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
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

Public Sub AU6371ELNoteBookSorting()
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
                 
                
                 
                 CardResult = DO_WritePort(card, Channel_P1A, &H7F) ' power on  for detect Normal mode
                 
                 Call MsecDelay(1.2)
                      
                      
                  If GetDeviceName(ChipString) <> "" Then
                    rv0 = 0
                    GoTo AU6371ELResult
                  
                  End If
                  
                 
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

Public Sub AU6371AFTest()

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
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
               
           ' 1. initail variable
                 ChipString = "vid_8751"
                
             
                    AU6371EL_SD = 0
                    AU6371EL_CF = 0
                    AU6371EL_XD = 0
                    AU6371EL_MS = 0
                    AU6371EL_MSP = 0
                    AU6371EL_BootTime = 0
           
            ' 2. set power
           
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
           
                CardResult = DO_ReadPort(card, Channel_P1B, LightOff)
                  
                If CardResult <> 0 Then
                    MsgBox "Read light off fail"
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
                     
                  '==== check light
                     
                     If rv0 <> 0 Then
                          If LightOn <> &HBF Or LightOff <> &HBF Then
                                    
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

                ' Because AU6371AFF has not CF test
                rv1 = 1  '-----------
            
                 Call LabelMenu(1, rv1, rv0)
            
                 Tester.Print rv1, " \\CF :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                '===============================================
                '  SMC Card test  : stop these test for card not enough
                '================================================
             
            
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
                   If Left(ChipName, 7) <> "AU6366C" And Left(ChipName, 10) <> "AU6371GLF2" And Left(OldChipName, 10) <> "AU6371HLF2" Then
                        CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                         Call MsecDelay(0.03)
                         
                        If CardResult <> 0 Then
                            MsgBox "Set MS Card Detect On Fail"
                            End
                        End If
                
                        If rv3 = 1 Then
                            CardResult = DO_WritePort(card, Channel_P1A, &H6F)
                        End If
                            Call MsecDelay(AU6371EL_BootTime)
                        If CardResult <> 0 Then
                           MsgBox "Set MS Card Detect Down Fail"
                           End
                        End If
                  If ChipName = "AU6371EL" Then
                   ReaderExist = 0
                 End If
                
                
                    ClosePipe
                    rv4 = CBWTest_New(0, rv3, ChipString)
                     ClosePipe
                    Call LabelMenu(3, rv4, rv3)
                
                Else
                     rv4 = 1
                End If
               
                     Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
                If ChipName <> "AU6371EL" Or OldChipName = "AU6371HLF2" Then
                
                
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
                Else
                
                rv5 = 1
                End If
                
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

Public Sub AU6371CFTest()
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
                
                Dim AU6371EL_SD As Byte
                Dim AU6371EL_CF As Byte
                Dim AU6371EL_XD As Byte
                Dim AU6371EL_MS As Byte
                Dim AU6371EL_MSP  As Byte
                Dim AU6371EL_BootTime As Single
                OldChipName = ""
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
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
                   
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
                 
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
                       Call LabelMenu(1, rv1, rv0)
                     ClosePipe
                 End If
              
                 If rv1 = 1 Then
                   
                    ClosePipe
                     rv1 = CBWTest_New_AU6375IncPattern2(0, 1, "vid_058f")
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
                   If Left(ChipName, 7) <> "AU6366C" And Left(ChipName, 10) <> "AU6371GLF2" And Left(OldChipName, 10) <> "AU6371HLF2" Then
                        CardResult = DO_WritePort(card, Channel_P1A, &H7F)
                         Call MsecDelay(0.03)
                         
                        If CardResult <> 0 Then
                            MsgBox "Set MS Card Detect On Fail"
                            End
                        End If
                
                        If rv3 = 1 Then
                            CardResult = DO_WritePort(card, Channel_P1A, &H6F)
                        End If
                            Call MsecDelay(AU6371EL_BootTime)
                        If CardResult <> 0 Then
                           MsgBox "Set MS Card Detect Down Fail"
                           End
                        End If
                  If ChipName = "AU6371EL" Then
                   ReaderExist = 0
                 End If
                
                
                    ClosePipe
                    rv4 = CBWTest_New(0, rv3, ChipString)
                     ClosePipe
                    Call LabelMenu(3, rv4, rv3)
                
                Else
                     rv4 = 1
                End If
               
                     Tester.Print rv4, " \\MS :0 Unknow device, 1 pass ,2 write fail, 3 read fail, 4 preious slot fail"
                   '===============================================
                '  MS Pro Card test
                '================================================
                If ChipName <> "AU6371EL" Or OldChipName = "AU6371HLF2" Then
                
                
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
                Else
                
                rv5 = 1
                End If
                
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

